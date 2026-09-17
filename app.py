import os
import re
import json
import time
import hashlib
import logging
import copy
from datetime import datetime, date, timezone
from io import BytesIO
import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Inches

st.set_page_config(page_title="AHJ Research Assistant v27.01", page_icon="🏛️", layout="wide")

# ============================================================
# CONFIGURATION & SECRETS
# ============================================================
STATE_OPTIONS = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "District of Columbia", "Florida", "Georgia",
    "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky",
    "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
    "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire",
    "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota",
    "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island",
    "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont",
    "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming",
]
PROJECT_TYPES = ["Replacement / Repair", "Remodel / Tenant Improvement", "Addition", "New Construction", "Site / Civil Work", "Other"]
BUILDING_CLASSES = ["Commercial", "Assembly", "Institutional", "Industrial", "Agricultural", "Residential (1-2 Family)", "Residential (Multi-family)", "Mixed-use", "Unknown"]

FACT_SOURCES = {"USER_PROVIDED", "RETRIEVED_RECORD", "AUTHORITATIVE_SOURCE", "INFERRED", "UNKNOWN"}

# 1. Add EVIDENCE_PROPOSITION_TYPES
EVIDENCE_PROPOSITION_TYPES = {
    "JURISDICTION",
    "AUTHORITY_HIERARCHY",
    "CODE_CURRENCY",
    "APPLICABILITY",
    "PERMIT_REQUIREMENT",
    "REVIEW_REQUIREMENT",
    "PERMIT_EXEMPTION",
    "PATHWAY",
    "THRESHOLD",
    "ENTITLEMENT",
    "OTHER",
}

GEMINI_KEY = os.getenv("GEMINI_KEY") or st.secrets.get("GEMINI_KEY", "")
PROMPT_VERSION = "v27.00_research_brief"

# ============================================================
# HELPERS & VALIDATION
# ============================================================
def discipline_family(value):
    value = str(value or "").strip().lower()
    # Composite display labels must resolve to the discipline family they
    # explicitly name.  Otherwise "Building / Structural" was classified as
    # building merely because "building" appeared first, while evidence
    # labeled "Structural" classified as structural.  That made valid
    # discipline-matched evidence look mismatched after validation repair.
    if "structural" in value:
        return "structural"
    if "accessibility" in value or "ada" in value:
        return "accessibility"
    families = {
        "mechanical": ["mechanical", "hvac", "heating", "cooling"],
        "electrical": ["electrical", "electric"],
        "structural": ["structural", "structure"],
        "energy": ["energy", "energy efficiency"],
        "planning": ["planning", "land use", "zoning", "cup"],
        "fire": ["fire", "life safety"],
        "building": ["building", "construction"],
    }
    for family, terms in families.items():
        if any(term in value for term in terms):
            return family
    return value

def _jurisdiction_process_source_mismatch(evidence, jurisdiction):
    """Return True when process evidence is tied to a governmental authority that
    conflicts with the dossier's verified site jurisdiction.

    A postal city is not the municipal AHJ for an unincorporated parcel. In the
    La Puente example, Gemini found a City of La Puente plan-check document for
    an address whose verified jurisdiction is unincorporated Los Angeles County.
    That source may be a real government document, but it cannot establish the
    County process for this parcel.
    """
    if not isinstance(evidence, dict) or not isinstance(jurisdiction, dict):
        return False
    city = _norm_text(jurisdiction.get("city"))
    county = _norm_text(jurisdiction.get("county"))
    ahj = _norm_text(jurisdiction.get("ahj"))
    if str(jurisdiction.get("status") or "").upper() != "VERIFIED":
        # An unresolved municipal boundary cannot authorize city-specific process evidence.
        # Postal-city identity is insufficient.
        if city and city != "not yet established":
            return True
        return True
    if city != "unincorporated":
        return False

    url = str(evidence.get("url") or "").strip().lower()
    title = _norm_text(evidence.get("title"))
    rule = _norm_text(evidence.get("rule"))
    authority = _norm_text(evidence.get("authority"))
    combined = " ".join([title, rule, authority, url])

    # Explicit municipal sources are not evidence of the county process for an
    # unincorporated parcel unless the source itself establishes that the county
    # uses/adopts that municipal process (which must be stated in the source).
    municipal_markers = [
        "city of ", "city department", "municipal", "city council",
        "city planning", "city building", "city development services",
    ]
    county_markers = [
        "los angeles county", "county of los angeles", "lacounty.gov",
        "dpw.lacounty.gov", "pw.lacounty.gov", "planning.lacounty.gov",
        "fire.lacounty.gov", "regional planning",
    ]
    if any(m in combined for m in municipal_markers):
        if not any(m in combined for m in county_markers):
            return True

    # Known municipal host for the current test parcel. This is deliberately
    # narrow rather than treating every city-related word as a mismatch.
    if "lapuente.org" in url:
        return True

    # A process rule that explicitly routes the work "to the City" conflicts
    # with an unincorporated County jurisdiction unless the same rule identifies
    # a County authority as the governing reviewer.
    if re.search(r"\b(?:submit|submitted|route|routed|apply|application|plan(?:s)?|review)\b.{0,120}\bto the city\b", rule, re.I):
        if not any(m in combined for m in county_markers):
            return True
    return False


def evidence_supports_permit_requirement(evidence, discipline):
    if not evidence:
        return False
    ev_disc = (evidence.get("discipline") or "").strip()
    current_disc = (discipline or "").strip()
    ev_family = discipline_family(ev_disc)
    current_family = discipline_family(current_disc)
    if ev_family != current_family:
        return False
    return evidence.get("proposition_type") == "PERMIT_REQUIREMENT" and not evidence_proposition_integrity_errors(evidence) and not evidence_source_specificity_errors([evidence])

def evidence_supports_pathway(evidence, discipline):
    if not evidence:
        return False
    ev_disc = (evidence.get("discipline") or "").strip()
    current_disc = (discipline or "").strip()
    ev_family = discipline_family(ev_disc)
    current_family = discipline_family(current_disc)
    if ev_family != current_family:
        return False
    return evidence.get("proposition_type") == "PATHWAY" and not evidence_proposition_integrity_errors(evidence) and not evidence_source_specificity_errors([evidence])

def evidence_supports_permit_exemption(evidence, discipline):
    if not evidence:
        return False
    ev_disc = (evidence.get("discipline") or "").strip()
    current_disc = (discipline or "").strip()
    if discipline_family(ev_disc) != discipline_family(current_disc):
        return False
    return evidence.get("proposition_type") == "PERMIT_EXEMPTION" and not evidence_proposition_integrity_errors(evidence) and not evidence_source_specificity_errors([evidence])

def evidence_supports_authority_hierarchy(evidence):
    if not evidence:
        return False
    return (
        evidence.get("proposition_type") == "AUTHORITY_HIERARCHY"
        and not evidence_proposition_integrity_errors(evidence)
        and bool(evidence.get("rule"))
    )

def evidence_supports_threshold(evidence, discipline):
    if not evidence:
        return False
    ev_disc = (evidence.get("discipline") or "").strip()
    current_disc = (discipline or "").strip()
    if discipline_family(ev_disc) != discipline_family(current_disc):
        return False
    return evidence.get("proposition_type") == "THRESHOLD" and not evidence_proposition_integrity_errors(evidence)

def _is_generic_review_source(evidence):
    """Reject generic agency/department pages as proof of a specific review process."""
    if not isinstance(evidence, dict):
        return True
    url = str(evidence.get("url") or "").strip()
    if not url or _is_grounding_redirect_url(url):
        return True
    try:
        from urllib.parse import urlparse
        path = (urlparse(url).path or "/").rstrip("/").lower() or "/"
    except Exception:
        return True
    generic = {
        "/", "/index.html", "/services", "/building-and-safety", "/bsd",
        "/building", "/planning", "/planning-and-zoning", "/zoning",
        "/permits", "/permit", "/permit-center", "/applications",
        "/application", "/fire", "/fire-safety", "/fire-prevention",
        "/public-works", "/publicworks", "/engineering", "/development",
        "/development-services", "/epicla", "/online-permits",
    }
    return path in generic or path.endswith(("/permit-portal", "/permitportal", "/online-permit"))

def evidence_supports_review(evidence, discipline):
    if not evidence:
        return False
    ev_disc = (evidence.get("discipline") or "").strip()
    current_disc = (discipline or "").strip()
    if discipline_family(ev_disc) != discipline_family(current_disc):
        return False
    return (
        evidence.get("proposition_type") == "REVIEW_REQUIREMENT"
        and not evidence_proposition_integrity_errors(evidence)
        and bool(str(evidence.get("url") or "").strip())
        and not _is_generic_review_source(evidence)
    )

def evidence_supports_any(evidence_ids, evidence_by_id, proposition_types, discipline):
    """Return True only when at least one referenced evidence item is valid for
    the requested proposition type(s), including source-specificity checks for
    propositions that require proposition-specific authority pages.

    Keep this helper aligned with evidence_ids_supporting_type() so validators
    and sanitizers cannot disagree about whether evidence is actually usable.
    """
    for proposition_type in proposition_types:
        if evidence_ids_supporting_type(
            evidence_ids, evidence_by_id, proposition_type, discipline
        ):
            return True
    return False

def extract_json(text):
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            return json.loads(text[start:end + 1])
        raise ValueError("No valid JSON found")

def _norm_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())

def evidence_internal_contradiction_errors(evidence):
    """Reject evidence whose generated findings explicitly negate its own proposition."""
    errors = []
    if not isinstance(evidence, dict):
        return errors
    eid = evidence.get("id", "Evidence")
    ptype = str(evidence.get("proposition_type") or "").upper()
    if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT"}:
        return errors

    # Gemini sometimes returns a strong conclusion plus a grounding/source finding
    # that says no proposition-specific rule was established.  Those two statements
    # cannot both be true.  Treat the contradiction as invalid evidence rather than
    # letting a later firewall decide which sentence to trust.
    contradiction_fields = (
        "source_finding", "rule_finding", "conclusion", "finding", "basis", "validation_note"
    )
    combined = " ".join(_norm_text(evidence.get(k)) for k in contradiction_fields)
    negations = [
        r"\bno\s+proposition[- ]specific\s+(?:authoritative\s+)?(?:rule|evidence)\s+was\s+established\b",
        r"\bno\s+proposition[- ]specific\s+(?:authoritative\s+)?rule\s+was\s+(?:found|identified|established)\b",
        r"\bproposition[- ]specific\s+(?:authoritative\s+)?(?:rule|evidence)\s+(?:was|were)\s+not\s+established\b",
        r"\b(?:no|not)\s+(?:explicit\s+)?(?:permit|approval|license)\s+(?:rule|requirement)\s+was\s+established\b",
    ]
    if any(re.search(p, combined, re.I) for p in negations):
        errors.append(
            f"{eid}: {ptype} evidence is internally contradictory; its generated findings state that no proposition-specific authoritative rule/evidence was established."
        )
    return errors

def evidence_proposition_integrity_errors(evidence):
    """Prevent evidence labels from being stronger than the extracted rule text."""
    errors = []
    if not evidence:
        return errors
    errors.extend(evidence_internal_contradiction_errors(evidence))
    eid = evidence.get("id", "Unknown")
    ptype = evidence.get("proposition_type")
    rule = _norm_text(evidence.get("rule"))
    if not rule:
        return errors

    permit_sentence_patterns = [
        re.compile(r"\b(?:permit|approval|license)\b[^.!?;:]{0,180}\b(?:is|are|be|become)?\s*(?:required|needed|necessary)\b"),
        re.compile(r"\b(?:required|needed|necessary|must|obtain|requires|require)\b[^.!?;:]{0,180}\b(?:permit|approval|license)\b"),
    ]
    exemption_explicit = re.compile(
        r"\bexempt(?:ed|ion)?\b|\bno\s+(?:[a-z -]+\s+)?permit\b|"
        r"\bpermit\s+(?:is\s+)?not\s+required\b|\bdoes\s+not\s+require\s+(?:a\s+)?permit\b|"
        r"\bnot\s+subject\s+to\s+(?:a\s+)?permit\b"
    )
    review_terms = re.compile(r"\breview\b|\binspection\b|\bsubmittal\b|\bcalculations?\b|\bplan review\b")
    pathway_terms = re.compile(r"\bsubmit\b|\bapplication\b|\bportal\b|\bonline\b|\bover[- ]the[- ]counter\b|\bprocess(?:ed|ing)?\b|\bfil(?:e|ing)\b|\bpermit center\b|\bplan review\b")
    threshold_terms = re.compile(r"\bthreshold\b|\bmore than\b|\bgreater than\b|\bless than\b|\bup to\b|\bover\s+\d|\bunder\s+\d|\bexceed(?:s|ing)?\b|\b(?:maximum|minimum)\b|\b\d+(?:\.\d+)?\s*(?:a|amp|amps|ampere|amperes|v|volt|volts|kv|kva|kw|va|w|watts?|lb|lbs|pounds?|sq\.?\s*ft|sf|cfm|btu|btuh|tons?|feet|ft|inches?|in\.)\b")

    if ptype == "PERMIT_REQUIREMENT" and not any(p.search(rule) for p in permit_sentence_patterns):
        errors.append(f"{eid}: PERMIT_REQUIREMENT label is unsupported by the evidence rule; the rule does not explicitly state a permit/approval/license requirement.")
    if ptype == "PERMIT_EXEMPTION" and not exemption_explicit.search(rule):
        errors.append(f"{eid}: PERMIT_EXEMPTION label is unsupported by the evidence rule; no explicit exemption/non-permit proposition is stated.")
    if ptype == "REVIEW_REQUIREMENT" and not review_terms.search(rule):
        errors.append(f"{eid}: REVIEW_REQUIREMENT label is unsupported; no review/inspection/submittal concept appears in the rule.")
    if ptype == "PATHWAY" and not pathway_terms.search(rule):
        errors.append(f"{eid}: PATHWAY label is unsupported; the rule does not describe a processing/submittal/review pathway.")
    if ptype == "THRESHOLD" and not threshold_terms.search(rule):
        errors.append(f"{eid}: THRESHOLD label is unsupported; no threshold/limit appears in the rule.")
    if ptype == "AUTHORITY_HIERARCHY":
        authority_terms = re.compile(
            r"\b(?:statewide|state code|state law|state standard|local amendment|local amendments|local code|local ordinance|"
            r"home rule|home-rule|preempt(?:ed|ion)?|local enforcement|delegat(?:ed|e) authority|"
            r"more stringent|less stringent|equivalent|minimum standard|shall prevail|prevails|controls|"
            r"adopted locally|locally adopted|statewide minimum)\b",
            re.I,
        )
        if not authority_terms.search(rule):
            errors.append(f"{eid}: AUTHORITY_HIERARCHY label is unsupported; the rule does not establish the relationship between state and local regulatory authority.")
    
    entitlement_terms = re.compile(
        r"\b(?:conditional use permit|CUP|land[- ]use approval|entitlement)\b"
        r".{0,220}\b(?:issued|approved|granted|amended|modified|modification|amendment|"
        r"required|not required|exempt|allowed|permitted|prohibited|condition|conditions|authorized|governs)\b",
        re.I,
    )
    if ptype == "ENTITLEMENT" and not entitlement_terms.search(rule):
        errors.append(f"{eid}: ENTITLEMENT label is unsupported; the rule does not establish a specific land-use entitlement, approval condition, amendment, or legal status.")
    
    return errors

def evidence_ids_supporting_type(evidence_ids, evidence_by_id, proposition_type, discipline):
    """Return evidence IDs that are actually valid for the requested proposition.

    IMPORTANT: proposition-specific source specificity is part of validity.
    Previously this helper checked proposition type/integrity but not source
    specificity, so a generic agency landing page could look like valid
    PERMIT_REQUIREMENT evidence to the deterministic firewall.  The semantic
    validator later rejected the same evidence, creating a loop where the
    firewall did not neutralize the unsupported permit finding.
    """
    result = []
    for eid in evidence_ids:
        evidence = evidence_by_id.get(eid)
        if not evidence:
            continue
        if discipline_family(evidence.get("discipline", "")) != discipline_family(discipline):
            continue
        if evidence.get("proposition_type") != proposition_type:
            continue
        if evidence_proposition_integrity_errors(evidence):
            continue

        # These proposition types require a proposition-specific source, not
        # merely an authoritative agency domain.  Keep the generic evidence
        # in the dossier's evidence trail, but do not count it as support.
        if proposition_type in {
            "PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT"
        } and evidence_source_specificity_errors([evidence]):
            continue

        result.append(eid)
    return result

def normalize_dossier_basis_values(data):
    if not isinstance(data, dict):
        return data
    basis_aliases = {
        "INSUFFICIENT_EVIDENCE": "NOT_ESTABLISHED",
        "INSUFFICIENT": "NOT_ESTABLISHED",
        "UNSUPPORTED": "NOT_ESTABLISHED",
        "NOT_SUPPORTED": "NOT_ESTABLISHED",
        "NO_EVIDENCE": "NOT_ESTABLISHED",
        "PARTIAL_EVIDENCE": "CONDITIONAL",
        "CONDITIONAL_EVIDENCE": "CONDITIONAL",
        "SUPPORTED": "DIRECT_EVIDENCE",
        "DIRECT": "DIRECT_EVIDENCE",
    }
    for item in data.get("disciplines", []) or []:
        if isinstance(item, dict):
            for key in ("permit_basis", "pathway_basis"):
                value = item.get(key)
                if isinstance(value, str):
                    item[key] = basis_aliases.get(value.strip().upper(), value)
    return data

def normalize_dossier_status_values(data):
    if not isinstance(data, dict):
        return data
    verified_aliases = {
        "VERIFIED": "VERIFIED_REQUIRED",
        "ESTABLISHED": "VERIFIED_REQUIRED",
    }
    uncertainty_aliases = {
        "NOT_ESTABLISHED": "CONDITIONAL",
        "INSUFFICIENT_EVIDENCE": "CONDITIONAL",
        "INSUFFICIENT": "CONDITIONAL",
        "UNSUPPORTED": "CONDITIONAL",
        "NOT_SUPPORTED": "CONDITIONAL",
        "NO_EVIDENCE": "CONDITIONAL",
        "UNKNOWN_EVIDENCE": "CONDITIONAL",
        "UNDETERMINED": "CONDITIONAL",
    }
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        for key in ("permit", "pathway"):
            value = item.get(key)
            if isinstance(value, str):
                normalized = re.sub(r"[^A-Z0-9_]+", "_", value.strip().upper()).strip("_")
                if normalized in verified_aliases:
                    item[key] = verified_aliases[normalized]
                    item.setdefault("validation_notes", []).append(f"{key} status '{value}' normalized to VERIFIED_REQUIRED.")
                elif normalized in uncertainty_aliases:
                    item[key] = uncertainty_aliases[normalized]
                    item.setdefault("validation_notes", []).append(f"{key} status '{value}' is not a valid status; normalized to CONDITIONAL. This does not establish a permit/pathway requirement.")
    return data

def has_definitive_negative_permit_claim(text):
    text = _norm_text(text)
    if not text:
        return False
    epistemic_markers = [
        "not established", "not determined", "cannot determine", "can't determine",
        "unable to determine", "unable to establish", "not established by current evidence",
        "not established by the current evidence", "not established by available evidence",
        "current evidence does not establish", "available evidence does not establish",
        "evidence does not establish", "evidence is insufficient", "insufficient evidence",
        "no evidence establishes", "no evidence currently establishes",
    ]
    if any(marker in text for marker in epistemic_markers):
        return False
    negative_patterns = [
        r"\bno\s+(?:separate\s+)?permit\s+(?:is\s+)?required\b",
        r"\bno\s+(?:separate\s+)?permit\s+is\s+needed\b",
        r"\bpermit\s+(?:is\s+)?not\s+required\b",
        r"\bpermit\s+(?:is\s+)?not\s+needed\b",
        r"\bdoes\s+not\s+require\s+(?:a\s+)?permit\b",
        r"\bdoes\s+not\s+trigger\s+(?:a\s+)?permit\b",
        r"\bno\s+permit\s+(?:is\s+)?necessary\b",
        r"\bexempt(?:ed|ion)?\s+from\s+(?:the\s+)?permit\b",
        r"\bexempt(?:ed|ion)?\s+from\s+permit\b",
        r"\bnot\s+subject\s+to\s+(?:a\s+)?permit\b",
    ]
    return any(re.search(pattern, text) for pattern in negative_patterns)

def is_epistemic_permit_finding(text):
    """Return True when permit wording explicitly says the requirement is unresolved.

    This is intentionally broader than a single phrase.  Deterministic permit
    validation must not mistake sentences such as "does not yet establish
    whether a permit is required" for an affirmative permit consequence.
    """
    text = _norm_text(text).lower()
    if not text:
        return False

    markers = [
        r"\bnot\s+(?:yet\s+)?established\b[^.]{0,220}\bpermit\b",
        r"\bnot\s+(?:yet\s+)?determined\b[^.]{0,220}\bpermit\b",
        r"\bcannot\s+determine\b[^.]{0,220}\bpermit\b",
        r"\bunable\s+to\s+(?:determine|establish)\b[^.]{0,220}\bpermit\b",
        r"\b(?:current|available)\s+evidence\s+does\s+not\s+(?:yet\s+)?establish\b[^.]{0,220}\bpermit\b",
        r"\bdoes\s+not\s+(?:by\s+itself\s+)?(?:yet\s+)?establish\b[^.]{0,220}\bpermit\b",
        r"\bwhether\b[^.]{0,120}\bpermit\s+(?:is\s+)?required\b",
    ]
    return any(re.search(pattern, text, re.I) for pattern in markers)


def semantic_consequence_errors(item, evidence_by_id):
    errors = []
    discipline = item.get("type", "Unknown")
    permit = item.get("permit")
    pathway = item.get("pathway")
    permit_basis = item.get("permit_basis")
    pathway_basis = item.get("pathway_basis")
    permit_text = _norm_text(item.get("permit_finding"))
    pathway_text = _norm_text(item.get("pathway_finding"))
    permit_ids = item.get("permit_evidence") or []
    pathway_ids = item.get("pathway_evidence") or []

    valid_permit = evidence_ids_supporting_type(permit_ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline)
    valid_exemption = evidence_ids_supporting_type(permit_ids, evidence_by_id, "PERMIT_EXEMPTION", discipline)
    valid_pathway = evidence_ids_supporting_type(pathway_ids, evidence_by_id, "PATHWAY", discipline)
    valid_review = evidence_ids_supporting_type(pathway_ids, evidence_by_id, "REVIEW_REQUIREMENT", discipline)

    epistemic_permit_finding = is_epistemic_permit_finding(permit_text)

    permit_consequence_patterns = [
        r"\bpermit\s+(?:is\s+)?required\s+(?:if|when|once|where|provided)",
        r"\bpermit\s+requirement\s+(?:depends|turns)\s+on",
        r"\b(?:requires?|triggers?|necessitates?)\s+(?:a\s+)?(?:separate\s+)?permit\b",
        r"\b(?:a\s+)?permit\s+(?:would|will)\s+be\s+required\b",
        r"\bmust\s+obtain\s+(?:a\s+)?permit\b",
        r"\b(?:permit|approval)\s+is\s+triggered\s+by\b",
    ]
    claims_conditional_permit = (
        False if epistemic_permit_finding
        else any(re.search(pattern, permit_text) for pattern in permit_consequence_patterns)
    )

    if claims_conditional_permit and not valid_permit and permit not in {"UNKNOWN", "NOT_ESTABLISHED"}:
        errors.append(f"{discipline}: permit finding states a conditional permit consequence without valid PERMIT_REQUIREMENT evidence.")

    if permit in {"VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED"} and permit_text:
        # IMPORTANT: epistemic non-establishment language is not a permit
        # consequence. Do not flag phrases such as "whether a permit is
        # required" or "does not establish whether a permit is required".
        if not epistemic_permit_finding and any(re.search(p, permit_text) for p in [
            r"\bpermit\s+(?:is\s+)?required\b",
            r"\bpermit\s+requirement\s+(?:depends|turns)\s+on",
            r"\b(?:requires?|triggers?)\s+(?:a\s+)?(?:separate\s+)?permit\b",
            r"\bmust\s+obtain\s+(?:a\s+)?permit\b",
        ]) and not valid_permit and not has_definitive_negative_permit_claim(permit_text):
            errors.append(f"{discipline}: permit conclusion is downstream of a rule but lacks direct PERMIT_REQUIREMENT evidence.")

    if permit == "VERIFIED_REQUIRED" and not valid_permit:
        errors.append(f"{discipline}: VERIFIED_REQUIRED permit requires valid PERMIT_REQUIREMENT evidence after semantic validation.")

    if has_definitive_negative_permit_claim(permit_text) and not valid_exemption:
        errors.append(f"{discipline}: definitive permit non-requirement/exemption lacks valid PERMIT_EXEMPTION evidence.")

    pathway_process_patterns = [
        r"\b(?:submit|file)\b",
        r"\b(?:through|via|using)\s+(?:the\s+)?(?:portal|permit\s+center|online)\b",
        r"\bprocessed\s+(?:as|through|via)\b",
        r"\bqualif(?:y|ies|ied)\s+for\b",
        r"\b(?:over[- ]the[- ]counter|trade\s+permit)\b",
        r"\bapplication\s+(?:is\s+)?(?:submitted|filed)\b",
        r"\bpathway\s+(?:is|would\s+be|depends)\b",
    ]
    claims_pathway = any(re.search(p, pathway_text) for p in pathway_process_patterns)
    if claims_pathway and pathway_basis == "DIRECT_EVIDENCE" and not valid_pathway:
        errors.append(f"{discipline}: DIRECT_EVIDENCE pathway finding lacks valid PATHWAY evidence.")
    if pathway == "VERIFIED_REQUIRED" and not valid_pathway:
        errors.append(f"{discipline}: VERIFIED_REQUIRED pathway requires valid PATHWAY evidence after semantic validation.")

    review_claim = any(re.search(p, pathway_text) for p in [
        r"\bplan review\b", r"\bengineering review\b", r"\breview\s+is\s+required\b", r"\binspection\s+is\s+required\b"
    ])
    if review_claim and pathway_basis == "DIRECT_EVIDENCE" and not (valid_pathway or valid_review):
        errors.append(f"{discipline}: direct pathway/review finding lacks PATHWAY or REVIEW_REQUIREMENT evidence.")

    return errors

def jurisdiction_evidence_integrity_errors(evidence):
    errors = []
    if not evidence:
        return errors
    eid = evidence.get("id", "Unknown")
    if evidence.get("proposition_type") != "JURISDICTION":
        return errors
    rule = _norm_text(evidence.get("rule"))
    if not rule:
        return errors
    jurisdiction_terms = re.compile(
        r"\b(?:located|lies|situated|within|inside|outside|unincorporated|incorporated|"
        r"municipal limits?|city limits?|county jurisdiction|jurisdiction|served by|"
        r"permitting authority|building authority|ahj)\b"
    )
    if not jurisdiction_terms.search(rule):
        errors.append(f"{eid}: JURISDICTION label is unsupported; the rule does not explicitly establish governmental jurisdiction.")
    return errors

def _jurisdiction_evidence_is_site_specific(evidence, address=""):
    if not isinstance(evidence, dict):
        return False
    text = _norm_text(" ".join([
        str(evidence.get("title") or ""),
        str(evidence.get("rule") or ""),
        str(evidence.get("retrieval_note") or ""),
    ]))
    addr = _norm_text(address)
    # Merely repeating the street address is NOT enough to establish municipal
    # jurisdiction. A municipal page, permit application, or property record
    # can mention a postal address while the parcel is actually unincorporated.
    # Require an explicit governmental boundary/jurisdiction result.
    #
    # This fixes the v26.30.52 failure where "13431 Temple Ave, La Puente, CA"
    # was treated as proof that the parcel was inside the City of La Puente.
    boundary_terms = [
        "unincorporated",
        "city limits",
        "municipal limits",
        "inside the city",
        "outside the city",
        "within the city",
        "within city limits",
        "outside city limits",
        "within municipal limits",
        "outside municipal limits",
        "jurisdiction is unincorporated",
        "jurisdiction: unincorporated",
        "jurisdiction is within",
        "jurisdiction is outside",
        "within the city of",
        "within city of",
        "inside the city of",
        "inside city limits",
        "outside the city of",
        "outside city limits",
        "within municipal limits",
        "outside municipal limits",
        "incorporated city of",
        "incorporated municipality",
        "unincorporated area",
        "unincorporated county",
        "county jurisdiction",
    ]
    return any(term in text for term in boundary_terms)

def _address_state_hint(address):
    """Return a conservative US state hint from an address.

    IMPORTANT: two-letter state abbreviations are only trusted when they occur
    in the address's state position (after the final comma, or as the final
    token before an optional ZIP code).  A broad substring search is unsafe:
    "La Puente, CA" contains the token "LA" in the city name and was falsely
    interpreted as Louisiana.
    """
    raw = str(address or "").strip()
    if not raw:
        return None

    state_map = {
        "alabama":"AL", "alaska":"AK", "arizona":"AZ", "arkansas":"AR",
        "california":"CA", "colorado":"CO", "connecticut":"CT", "delaware":"DE",
        "florida":"FL", "georgia":"GA", "hawaii":"HI", "idaho":"ID",
        "illinois":"IL", "indiana":"IN", "iowa":"IA", "kansas":"KS",
        "kentucky":"KY", "louisiana":"LA", "maine":"ME", "maryland":"MD",
        "massachusetts":"MA", "michigan":"MI", "minnesota":"MN", "mississippi":"MS",
        "missouri":"MO", "montana":"MT", "nebraska":"NE", "nevada":"NV",
        "new hampshire":"NH", "new jersey":"NJ", "new mexico":"NM", "new york":"NY",
        "north carolina":"NC", "north dakota":"ND", "ohio":"OH", "oklahoma":"OK",
        "oregon":"OR", "pennsylvania":"PA", "rhode island":"RI", "south carolina":"SC",
        "south dakota":"SD", "tennessee":"TN", "texas":"TX", "utah":"UT",
        "vermont":"VT", "virginia":"VA", "washington":"WA", "west virginia":"WV",
        "wisconsin":"WI", "wyoming":"WY", "district of columbia":"DC",
    }

    # Full state names are safe to recognize as names, but do so on word
    # boundaries rather than substring matching.
    text = _norm_text(raw)
    for name, abbr in state_map.items():
        if re.search(rf"(?:^|[,\s]){re.escape(name)}(?:$|[,\s])", text):
            return abbr

    valid_abbr = set(state_map.values())

    # Most normal addresses are "City, ST" or "City, ST ZIP".  Only inspect
    # the final comma-delimited component so city names such as "La Puente"
    # cannot be mistaken for Louisiana (LA).
    tail = re.split(r",", text)[-1].strip()
    tail = re.sub(r"\b(?:usa|united states)\b", " ", tail).strip()
    m = re.match(r"^([a-z]{2})(?:\s+\d{5}(?:-\d{4})?)?\s*$", tail)
    if m and m.group(1).upper() in valid_abbr:
        return m.group(1).upper()

    # Also support addresses without commas, e.g. "123 Main St Portland OR
    # 97201", while requiring the abbreviation to be immediately before an
    # optional ZIP/end-of-string. This prevents matching ordinary words.
    m = re.search(r"(?:^|\s)([a-z]{2})(?:\s+\d{5}(?:-\d{4})?)?\s*$", text)
    if m and m.group(1).upper() in valid_abbr:
        return m.group(1).upper()

    return None

def _normalize_state_code(value):
    """Normalize a state selector/name to its USPS-style two-letter code."""
    raw = _norm_text(value).strip().lower()
    if not raw:
        return None
    state_map = {
        "alabama":"AL", "alaska":"AK", "arizona":"AZ", "arkansas":"AR",
        "california":"CA", "colorado":"CO", "connecticut":"CT", "delaware":"DE",
        "florida":"FL", "georgia":"GA", "hawaii":"HI", "idaho":"ID",
        "illinois":"IL", "indiana":"IN", "iowa":"IA", "kansas":"KS",
        "kentucky":"KY", "louisiana":"LA", "maine":"ME", "maryland":"MD",
        "massachusetts":"MA", "michigan":"MI", "minnesota":"MN", "mississippi":"MS",
        "missouri":"MO", "montana":"MT", "nebraska":"NE", "nevada":"NV",
        "new hampshire":"NH", "new jersey":"NJ", "new mexico":"NM", "new york":"NY",
        "north carolina":"NC", "north dakota":"ND", "ohio":"OH", "oklahoma":"OK",
        "oregon":"OR", "pennsylvania":"PA", "rhode island":"RI", "south carolina":"SC",
        "south dakota":"SD", "tennessee":"TN", "texas":"TX", "utah":"UT",
        "vermont":"VT", "virginia":"VA", "washington":"WA", "west virginia":"WV",
        "wisconsin":"WI", "wyoming":"WY", "district of columbia":"DC",
    }
    if len(raw) == 2 and raw.upper() in set(state_map.values()):
        return raw.upper()
    return state_map.get(raw)

def validate_input_state_consistency(data, address, selected_state):
    """Catch an obvious address/state selector contradiction without changing it.

    State names (e.g. California) and USPS abbreviations (CA) are equivalent;
    compare normalized codes so the validator does not create a false failure.
    """
    addr_state = _address_state_hint(address)
    selected_code = _normalize_state_code(selected_state)
    if addr_state and selected_code and addr_state != selected_code:
        return [
            f"Input contradiction: address appears to be in {addr_state}, but the selected state/jurisdiction is {selected_state}. "
            "Correct the project metadata before relying on the dossier."
        ]
    return []

def sanitize_unverifiable_verified_jurisdiction(data, address=""):
    if not isinstance(data, dict):
        return data
    jurisdiction = data.get("jurisdiction") or {}
    if jurisdiction.get("status") != "VERIFIED":
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    ids = jurisdiction.get("evidence") or []
    valid = [
        evidence_by_id[eid] for eid in ids
        if eid in evidence_by_id
        and evidence_by_id[eid].get("proposition_type") == "JURISDICTION"
        and not evidence_proposition_integrity_errors(evidence_by_id[eid])
        and not jurisdiction_evidence_integrity_errors(evidence_by_id[eid])
    ]
    site_specific = [e for e in valid if _jurisdiction_evidence_is_site_specific(e, address)]
    if site_specific:
        combined = _norm_text(" ".join(str(e.get("rule") or "") for e in site_specific))
        if "unincorporated" in combined:
            city = _norm_text(jurisdiction.get("city"))
            if city and city not in {"unincorporated", "unincorporated area", "unincorporated county"}:
                jurisdiction["city"] = "Unincorporated"
        return data
    jurisdiction["status"] = "CONDITIONAL"
    jurisdiction["ahj"] = jurisdiction.get("ahj") or "Unconfirmed"
    jurisdiction["evidence"] = []
    jurisdiction["validation_note"] = (
        "Jurisdiction downgraded because the cited evidence did not establish the "
        "project parcel's actual governmental jurisdiction and municipal boundary result. "
        "Postal city/ZIP, generic agency coverage, parcel references, and map/GIS mentions "
        "do not by themselves establish municipal boundaries."
    )
    return data

def sanitize_invalid_authority_evidence_links(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        raw = item.get("authority_evidence") or []
        if isinstance(raw, str):
            raw = [raw]
        if not isinstance(raw, list):
            raw = []
        valid = [
            eid for eid in raw
            if eid in evidence_by_id and evidence_supports_authority_hierarchy(evidence_by_id[eid])
        ]
        if raw != valid:
            item["authority_evidence"] = valid
            item.setdefault("validation_notes", []).append(
                "Removed authority_evidence links that were not AUTHORITY_HIERARCHY evidence; "
                "the underlying source records were preserved."
            )
    return data

# NEW SANITIZER ADDED HERE
def sanitize_cross_discipline_applicability_links(data):
    """
    Prevent clearly unrelated evidence from being treated as direct
    applicability evidence.

    Evidence discipline labels are metadata, not proof that the source
    cannot establish applicability for another discipline. General Building
    and Administrative sources may legitimately establish applicability
    across disciplines when the source rule itself applies broadly.

    Only clearly unrelated evidence is downgraded. The evidence record is
    preserved for traceability.
    """
    if not isinstance(data, dict):
        return data

    evidence_by_id = {
        e.get("id"): e
        for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    cross_discipline_families = {
        "administrative",
        "general_building",
        "building",
        "building_structural",
        "general",
    }

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue

        applicability = item.get("applicability")
        if not isinstance(applicability, dict):
            continue

        relationship = str(
            applicability.get("relationship") or ""
        ).strip().lower()

        if relationship != "direct":
            continue

        discipline = str(
            item.get("type")
            or item.get("discipline")
            or "Unknown"
        )

        current_family = discipline_family(discipline)

        raw_ids = applicability.get("evidence") or []
        if isinstance(raw_ids, str):
            raw_ids = [raw_ids]
        if not isinstance(raw_ids, list):
            continue

        bad_ids = []

        for eid in raw_ids:
            evidence = evidence_by_id.get(eid)
            if not evidence:
                continue

            evidence_family = discipline_family(
                str(evidence.get("discipline") or "")
            )

            if not evidence_family or not current_family:
                continue

            if evidence_family == current_family:
                continue

            # General Building / Administrative evidence can legitimately
            # establish cross-discipline applicability. The source rule, not
            # the metadata label alone, determines substantive relevance.
            if evidence_family in cross_discipline_families:
                continue

            bad_ids.append(eid)

        if bad_ids:
            applicability["relationship"] = "conditional"
            notes = item.setdefault("validation_notes", [])
            notes.append(
                "Applicability was downgraded from direct because "
                f"the cited evidence ({', '.join(bad_ids)}) appears to belong "
                "to an unrelated discipline. The evidence remains available "
                "in the research trail."
            )

    return data

def sanitize_unsupported_entitlement_rules(data):
    if not isinstance(data, dict):
        return data
    raw_evidence = data.get("evidence")
    evidence_list = raw_evidence if isinstance(raw_evidence, list) else []
    evidence_by_id = {
        e.get("id"): e for e in evidence_list
        if isinstance(e, dict) and e.get("id")
    }
    disciplines = data.get("disciplines")
    if not isinstance(disciplines, list):
        return data
    for item in disciplines:
        if not isinstance(item, dict):
            continue
        if discipline_family(item.get("type")) != "planning":
            continue
        app = item.get("applicability")
        if not isinstance(app, dict):
            continue
        rule = _norm_text(app.get("rule"))
        ids = app.get("evidence")
        if not isinstance(ids, list):
            ids = []
        valid_entitlement = evidence_ids_supporting_type(
            ids, evidence_by_id, "ENTITLEMENT", item.get("type", "Planning / Land Use")
        )
        if valid_entitlement or not rule:
            continue
        unsupported_consequence = (
            re.search(
                r"\b(?:CUP|conditional use permit|land[- ]use approval|entitlement)\b.{0,220}"
                r"\b(?:amendment|modification|approval|clearance)\b.{0,100}"
                r"\b(?:required|not required|needed|not needed|necessary|unnecessary|no)\b",
                rule, re.I,
            )
            or re.search(
                r"\b(?:requires? no|does not require|doesn't require|no)\b.{0,100}"
                r"\b(?:CUP|conditional use permit|amendment|modification)\b",
                rule, re.I,
            )
        )
        if unsupported_consequence:
            app["rule"] = (
                "Land-use regulations may govern the project, but the existing entitlement "
                "conditions and any amendment consequence are not established by the current evidence."
            )
            app["determination"] = "cannot_determine"
            app["relationship"] = "conditional"
            app["validation_note"] = (
                "Planning applicability rule was neutralized because it contained a CUP/"
                "entitlement consequence without discipline-matched ENTITLEMENT evidence."
            )
            item["applicability"] = app
    return data

def sanitize_unsupported_entitlement_conclusions(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {e.get("id"): e for e in (data.get("evidence") or []) if isinstance(e, dict) and e.get("id")}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict) or discipline_family(str(item.get("type") or "")) != "planning":
            continue
        discipline = item.get("type", "Planning / Land Use")
        ids = item.get("permit_evidence") or []
        valid_entitlement = evidence_ids_supporting_type(ids, evidence_by_id, "ENTITLEMENT", discipline)
        finding = _norm_text(item.get("permit_finding"))
        consequence_patterns = [
            r"\bCUP\b.{0,180}\b(?:amendment|modification)\b",
            r"\b(?:CUP|land[- ]use|planning)\b.{0,180}\b(?:approval|clearance)\b.{0,100}\b(?:required|needed|necessary)\b",
            r"\b(?:land[- ]use|planning)\b.{0,180}\b(?:permit|approval)\b.{0,100}\b(?:required|needed|necessary)\b",
        ]
        epistemic = any(re.search(p, finding, re.I) for p in [r"\bnot established\b", r"\bcannot determine\b", r"\bcurrent evidence does not establish\b", r"\bdoes not by itself establish\b"])
        if any(re.search(p, finding, re.I) for p in consequence_patterns) and not valid_entitlement and not epistemic:
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                "A planning / land use approval or CUP amendment requirement is not established by current entitlement evidence. "
                "Existing CUP conditions may affect the outcome, but project facts alone do not establish the legal consequence. "
                "Confirm the applicable requirement using the existing approval and an authoritative planning source."
            )
            pathway = _norm_text(item.get("pathway_finding"))
            pathway_ids = item.get("pathway_evidence") or []
            valid_pathway = (
                evidence_ids_supporting_type(pathway_ids, evidence_by_id, "PATHWAY", discipline)
                or evidence_ids_supporting_type(pathway_ids, evidence_by_id, "REVIEW_REQUIREMENT", discipline)
            )
            pathway_claim = any(re.search(p, pathway, re.I) for p in [
                r"\bplanning clearance\b.{0,120}\b(?:verified|required|must|necessary)\b",
                r"\bCUP\b.{0,120}\b(?:clearance|approval|review)\b.{0,100}\b(?:required|verified|must)\b",
            ])
            if pathway_claim and not valid_pathway and not re.search(r"\b(?:not established|cannot determine)\b", pathway, re.I):
                item["pathway"] = "CONDITIONAL"
                item["pathway_basis"] = "NOT_ESTABLISHED"
                item["pathway_evidence"] = []
                item["pathway_finding"] = (
                    "The planning / land use processing pathway is not established by current evidence. "
                    "Confirm whether existing CUP conditions require planning review before or with the mechanical permit."
                )
    return data

def sanitize_unverifiable_verified_pathways(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    changed = []
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict) or item.get("pathway") != "VERIFIED_REQUIRED":
            continue
        discipline = str(item.get("type") or "Unknown")
        ids = item.get("pathway_evidence") or []
        valid = evidence_ids_supporting_type(ids, evidence_by_id, "PATHWAY", discipline)
        if valid:
            continue
        item["pathway"] = "CONDITIONAL"
        item["pathway_basis"] = "NOT_ESTABLISHED"
        item["pathway_evidence"] = []
        item["pathway_finding"] = (
            f"The {discipline.lower()} processing pathway is not established by "
            "current evidence. Confirm the applicable submission or review process "
            "with an authoritative source."
        )
        changed.append(discipline)
    if not changed:
        return data
    bottom = _norm_text(data.get("bottom_line"))
    if bottom:
        for discipline in changed:
            d = re.escape(discipline.lower())
            bottom = re.sub(
                rf"\b{d}\b[^.]*\b(?:pathway|submission|plan review|portal)\b[^.]*\b(?:required|must|requires?|established)\b[^.]*\.",
                "", bottom, flags=re.I
            )
        data["bottom_line"] = re.sub(r"\s{2,}", " ", bottom).strip()
    return data

def sanitize_unsupported_pathway_conclusions(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    process_patterns = [
        r"\bsubmit(?:ted|s|ting)?\b", r"\bapplication(?:s)?\b", r"\bportal\b",
        r"\bprocessed\b", r"\bfile(?:d|s|ing)?\b", r"\bover[- ]the[- ]counter\b",
        r"\bplan review\b", r"\bpermit center\b", r"\belectronic(?:ally)?\b",
        r"\bconcurrently\b", r"\bseparately\b",
    ]
    epistemic_patterns = [
        r"\bnot established\b", r"\bcannot determine\b", r"\bunable to determine\b",
        r"\bcurrent evidence does not establish\b", r"\bnot currently established\b",
    ]
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        text = _norm_text(item.get("pathway_finding"))
        if not text or not any(re.search(p, text) for p in process_patterns):
            continue
        if any(re.search(p, text) for p in epistemic_patterns):
            continue
        ids = item.get("pathway_evidence") or []
        valid = evidence_ids_supporting_type(ids, evidence_by_id, "PATHWAY", discipline) or evidence_ids_supporting_type(ids, evidence_by_id, "REVIEW_REQUIREMENT", discipline)
        if valid:
            continue
        item["pathway"] = "CONDITIONAL"
        item["pathway_basis"] = "NOT_ESTABLISHED"
        item["pathway_evidence"] = []
        item["pathway_finding"] = (
            f"The {str(discipline).lower()} processing pathway is not established by current evidence. "
            "Confirm the applicable submission or review process with an authoritative source."
        )
    return data

def sanitize_invalid_evidence_propositions(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    quarantined = set()
    for eid, evidence in evidence_by_id.items():
        integrity_errors = evidence_proposition_integrity_errors(evidence)
        source_errors = []
        original = evidence.get("proposition_type")
        if original in {
            "PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT"
        }:
            # Use the exact same source-specificity contract as the validator.
            # This is intentionally broader than the rule-text integrity check:
            # a specific-looking URL with a weak rule, or a generic portal URL,
            # must not survive as proposition-specific evidence.
            source_errors = evidence_source_specificity_errors([evidence])

        if (integrity_errors or source_errors) and original in EVIDENCE_PROPOSITION_TYPES and original != "OTHER":
            evidence["proposition_type"] = "OTHER"
            reasons = []
            if integrity_errors:
                reasons.append("the source rule did not explicitly support the declared proposition")
            if source_errors:
                reasons.append("the source did not satisfy the proposition-specificity contract")
            evidence["validation_note"] = (
                "Quarantined: " + "; ".join(reasons) +
                ". The source record is preserved for review."
            )
            quarantined.add(eid)
    if not quarantined:
        return data
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        for field in ("permit_evidence", "pathway_evidence"):
            item[field] = [eid for eid in (item.get(field) or []) if eid not in quarantined]
        app = item.get("applicability") or {}
        if isinstance(app, dict):
            app["evidence"] = [eid for eid in (app.get("evidence") or []) if eid not in quarantined]
    data["bottom_line_evidence"] = [
        eid for eid in (data.get("bottom_line_evidence") or []) if eid not in quarantined
    ]
    return data

def has_concrete_threshold_claim(text, discipline=None):
    text = _norm_text(text)
    if not text:
        return False
    epistemic = re.compile(
        r"(?:not established|not determined|cannot determine|unable to determine|"
        r"current evidence does not establish|insufficient evidence|"
        r"specific .* threshold .* not established|threshold .* not established|"
        r"no .* threshold .* established)"
    )
    if epistemic.search(text):
        return False
    general_patterns = [
        r"\bover\s+\d+(?:\.\d+)?",
        r"\bunder\s+\d+(?:\.\d+)?",
        r"\bmore than\s+\d+(?:\.\d+)?",
        r"\bless than\s+\d+(?:\.\d+)?",
        r"\bgreater than\s+\d+(?:\.\d+)?",
        r"\bexceeds?\s+\d+(?:\.\d+)?",
        r"\bup to\s+\d+(?:\.\d+)?",
        r"\b(?:maximum|minimum)\s+(?:of\s+)?\d+(?:\.\d+)?",
        r"\b\d+(?:\.\d+)?\s*(?:lb|lbs|pounds?|cfm|btu|btuh|tons?|sq\.?\s*ft|square feet|sf|kw|kva|va|w|watts?|volts?|kv|feet|ft|inches?|in\.)\b",
    ]
    if any(re.search(pattern, text) for pattern in general_patterns):
        return True
    if discipline_family(discipline or "") == "electrical":
        electrical_patterns = [
            r"\b\d+(?:\.\d+)?\s*a\b", r"\b\d+(?:\.\d+)?\s*amps?\b",
            r"\b\d+(?:\.\d+)?\s*v\b", r"\b\d+(?:\.\d+)?\s*volts?\b",
            r"\b\d+(?:\.\d+)?\s*kva\b", r"\b\d+(?:\.\d+)?\s*va\b",
            r"\b(?:mca|mop|ampacity|circuit rating|service size|overcurrent rating)\s*(?:of|=|:)?\s*(?:at least|minimum|maximum|up to)?\s*\d+(?:\.\d+)?\s*(?:a|amps?|ampere|amperes|v|volts?|kva|va|kw|w)\b",
            r"\b(?:minimum|maximum)\s+\d+(?:\.\d+)?\s*(?:a|amps?|ampere|amperes|v|volts?|kva|va|kw|w)\b",
            r"\b\d+(?:\.\d+)?\s*(?:a|amps?|ampere|amperes|v|volts?|kva|va|kw|w)\s+(?:mca|mop|ampacity|circuit rating|service size|overcurrent rating)\b",
        ]
        return any(re.search(pattern, text) for pattern in electrical_patterns)
    return False

def repair_missing_threshold_links(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        finding_text = " ".join([
            str(item.get("permit_finding") or ""),
            str(item.get("pathway_finding") or "")
        ])
        if not has_concrete_threshold_claim(finding_text, discipline):
            continue
        linked_ids = set(
            (item.get("permit_evidence") or []) +
            (item.get("pathway_evidence") or []) +
            ((item.get("applicability") or {}).get("evidence") or [])
        )
        if any(evidence_supports_threshold(evidence_by_id.get(eid), discipline) for eid in linked_ids):
            continue
        candidates = [
            eid for eid, evidence in evidence_by_id.items()
            if evidence_supports_threshold(evidence, discipline) and eid not in linked_ids
        ]
        if candidates:
            item["pathway_evidence"] = list(item.get("pathway_evidence") or [])
            item["pathway_evidence"].extend(candidates)
    return data

def sanitize_unsupported_threshold_conclusions(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        linked_ids = ((item.get("permit_evidence") or []) +
                      (item.get("pathway_evidence") or []) +
                      ((item.get("applicability") or {}).get("evidence") or []))
        has_threshold = any(
            evidence_supports_threshold(evidence_by_id.get(eid), discipline)
            for eid in linked_ids
        )
        if has_threshold:
            continue
        for field in ("permit_finding", "pathway_finding"):
            original = str(item.get(field) or "")
            if not has_concrete_threshold_claim(original, discipline):
                continue
            item[field] = (
                f"The specific {discipline.lower()} threshold or limit is not established by current evidence. "
                "Do not apply a numeric or categorical trigger until an authoritative, discipline-matched "
                "source establishes it."
            )
            if field == "permit_finding" and item.get("permit") == "VERIFIED_REQUIRED":
                item["permit"] = "CONDITIONAL"
                item["permit_basis"] = "NOT_ESTABLISHED"
                item["permit_evidence"] = []
    return data

def sanitize_unsupported_permit_conclusions(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        finding = _norm_text(item.get("permit_finding"))
        permit_ids = item.get("permit_evidence") or []
        valid_permit = evidence_ids_supporting_type(
            permit_ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline
        )
        valid_exemption = evidence_ids_supporting_type(
            permit_ids, evidence_by_id, "PERMIT_EXEMPTION", discipline
        )
        if valid_permit or valid_exemption:
            continue
        permit_claim_patterns = [
            r"\bpermit\s+(?:is\s+)?required\b",
            r"\bpermit\s+requirement\s+(?:depends|turns)\s+on",
            r"\b(?:requires?|triggers?|necessitates?)\s+(?:a\s+)?(?:separate\s+)?permit\b",
            r"\b(?:a\s+)?permit\s+(?:would|will|may|might|could)\s+be\s+required\b",
            r"\b(?:permit|approval)\s+(?:may|might|could|would)\s+be\s+(?:triggered|required|necessary)\b",
            r"\b(?:permit|approval)\s+(?:depends|turns)\s+on\b",
            r"\b(?:permit|approval)\s+(?:consequence|trigger)\b",
            r"\bmust\s+obtain\s+(?:a\s+)?permit\b",
            r"\b(?:permit|approval)\s+is\s+triggered\s+by\b",
        ]
        epistemic_nonclaim = any(re.search(pattern, finding) for pattern in [
            r"\bpermit\s+(?:requirement\s+)?(?:is\s+)?not\s+established\b",
            r"\bnot\s+established\b[^.]{0,180}\bpermit\b",
            r"\bcannot\s+determine\b[^.]{0,180}\bpermit\b",
            r"\bcan(?:not|'t)\s+be\s+determined\b[^.]{0,180}\bpermit\b",
            r"\bdoes\s+not\s+(?:by\s+itself\s+)?establish\b[^.]{0,180}\bpermit\b",
            r"\bnot\s+establish(?:ed|ing)?\b[^.]{0,180}\bpermit\b",
            r"\bwhether\s+(?:a\s+)?(?:[^.]{0,80}\s+)?permit\s+is\s+required\b",
        ])
        if epistemic_nonclaim:
            continue
        has_claim = any(re.search(pattern, finding) for pattern in permit_claim_patterns)
        if not has_claim and re.search(r"\bpermit\b", finding):
            conditional_terms = [
                r"\bmay\b", r"\bmight\b", r"\bcould\b", r"\bwould\b",
                r"\bdepends?\b", r"\bdepending\s+on\b", r"\bif\b", r"\bwhen\b",
                r"\btrigger(?:s|ed)?\b", r"\brequire(?:s|d)?\b", r"\bconsequence\b",
            ]
            has_claim = any(re.search(pattern, finding) for pattern in conditional_terms)
        if not has_claim:
            continue
        if has_definitive_negative_permit_claim(finding):
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"A {discipline.lower()} permit exemption or non-requirement is not established by current "
                "evidence. Confirm the exemption or permit requirement using an authoritative permit-specific source."
            )
            continue
        item["permit"] = "CONDITIONAL"
        item["permit_basis"] = "NOT_ESTABLISHED"
        item["permit_evidence"] = []
        item["permit_finding"] = (
            f"The available evidence does not yet establish whether a {discipline.lower()} permit is required. "
            "Once the missing project facts are confirmed, check the applicable permit-specific requirement."
        )
    broad_permit_claims = [
        r"\b(?:permit|trade permit|building permit)\b[^.!?]{0,220}\b(?:required|requires|trigger(?:s|ed)?|necessitat(?:es|ed)|must obtain|would be required|will be required)\b",
        r"\b(?:required|requires|trigger(?:s|ed)?|necessitat(?:es|ed)|must obtain|would be required|will be required)\b[^.!?]{0,220}\b(?:permit|trade permit|building permit)\b",
    ]
    epistemic_only = [
        r"\bpermit requirement\b[^.!?]{0,160}\bnot established\b",
        r"\bnot established\b[^.!?]{0,160}\bpermit(?: requirement)?\b",
        r"\bcannot determine\b[^.!?]{0,160}\bpermit(?: requirement)?\b",
        r"\bcurrent evidence\b[^.!?]{0,180}\bdoes not establish\b[^.!?]{0,80}\bpermit\b",
    ]
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        finding = _norm_text(item.get("permit_finding"))
        if not finding:
            continue
        valid = evidence_ids_supporting_type(
            item.get("permit_evidence") or [], evidence_by_id, "PERMIT_REQUIREMENT", discipline
        )
        if valid or has_definitive_negative_permit_claim(finding):
            continue
        # A non-establishment statement is intentionally allowed without
        # PERMIT_REQUIREMENT evidence; it is not a legal permit conclusion.
        if any(re.search(p, finding) for p in epistemic_only) or re.search(
            r"\bwhether\s+(?:a\s+)?(?:[^.]{0,80}\s+)?permit\s+is\s+required\b",
            finding,
        ):
            continue
        if any(re.search(p, finding) for p in broad_permit_claims):
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"The available evidence does not yet establish whether a {str(discipline).lower()} permit is required. "
                "Check the applicable discipline-specific permit requirement against an authoritative source."
            )
    return data

def _extract_years(text):
    return {int(y) for y in re.findall(r"\b(?:19|20)\d{2}\b", str(text or ""))}

def _extract_edition_years(text):
    text = str(text or "")
    years = set()
    pattern = re.compile(
        r"\b((?:19|20)\d{2})\s+(?:(?:Oregon|[A-Z][A-Za-z&/-]+)\s+){0,8}"
        r"(?:Code|Specialty\s+Code|OSSC|OMSC|OESC|OEESC)\b",
        re.I,
    )
    for match in pattern.finditer(text):
        years.add(int(match.group(1)))
    return years

def _has_future_effective_or_mandatory_date(text, as_of_date):
    if not as_of_date:
        return False
    month_map = {
        "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
        "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
    }
    pattern = re.compile(
        r"\b(?:effective|mandatory|in effect)\s*[:]?\s*"
        r"(?:[A-Za-z]+\s+)?(?:(?:1st|2nd|3rd|4th|5th|6th|7th|8th|9th|10th|11th|12th|13th|14th|15th|16th|17th|18th|19th|20th|21st|22nd|23rd|24th|25th|26th|27th|28th|29th|30th|31st|\d{1,2})[\s.-]+)?"
        r"([A-Za-z]+)\s+\d{1,2},?\s+((?:19|20)\d{2})", re.I)
    for m in pattern.finditer(str(text or "")):
        month = month_map.get(m.group(1).lower())
        if not month:
            continue
        tail = m.group(0)
        dm = re.search(r"(?:[A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})", tail)
        if not dm:
            continue
        try:
            from datetime import date as _date
            d = _date(int(dm.group(2)), month, int(dm.group(1)))
            if d > as_of_date:
                return True
        except Exception:
            pass
    return False

def code_currency_integrity_errors(code, evidence_by_id, as_of_date=None):
    errors = []
    if not isinstance(code, dict):
        return errors
    if code.get("status") != "CURRENT":
        return errors
    name = str(code.get("name") or "")
    ids = code.get("evidence") or []
    if not ids:
        return [f"Current code '{name}' must have CODE_CURRENCY evidence."]
    valid = []
    code_years = _extract_edition_years(name)
    for eid in ids:
        ev = evidence_by_id.get(eid)
        if not ev or ev.get("proposition_type") != "CODE_CURRENCY":
            continue
        if evidence_proposition_integrity_errors(ev):
            continue
        rule = _norm_text(ev.get("rule"))
        title = _norm_text(ev.get("title"))
        combined = f"{title} {rule}"
        if not re.search(r"\b(?:current|currently|adopted|adoption|effective|mandatory|in effect|phase[- ]in|latest|most recent)\b", combined):
            continue
        valid.append(ev)
    if not valid:
        errors.append(f"Current code '{name}' lacks valid CODE_CURRENCY evidence establishing current/adopted/effective status.")
        return errors
    evidence_years = set()
    for ev in valid:
        evidence_years |= _extract_edition_years(f"{ev.get('title','')} {ev.get('rule','')}")
    if code_years and evidence_years and not (code_years & evidence_years):
        errors.append(
            f"Current code '{name}' edition/year conflicts with its cited CODE_CURRENCY evidence "
            f"({sorted(evidence_years)})."
        )
    if as_of_date:
        future_years = {y for y in evidence_years if y > as_of_date.year}
        if future_years and not any(y <= as_of_date.year for y in evidence_years):
            errors.append(
                f"Current code '{name}' is supported only by future-dated code evidence "
                f"relative to {as_of_date.isoformat()}."
            )
        if _has_future_effective_or_mandatory_date(" ".join(
            f"{ev.get('title','')} {ev.get('rule','')}" for ev in valid
        ), as_of_date):
            errors.append(
                f"Current code '{name}' has an effective/mandatory date after the project date "
                f"{as_of_date.isoformat()}."
            )
    return errors

def sanitize_invalid_current_codes(data, as_of_date=None):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for code in data.get("codes", []):
        if not isinstance(code, dict) or code.get("status") != "CURRENT":
            continue
        ids = code.get("evidence") or []
        currency_evidence = [
            evidence_by_id[eid] for eid in ids
            if eid in evidence_by_id
            and evidence_by_id[eid].get("proposition_type") == "CODE_CURRENCY"
            and not evidence_proposition_integrity_errors(evidence_by_id[eid])
        ]
        if currency_evidence and as_of_date:
            supported_years = set()
            for ev in currency_evidence:
                supported_years |= _extract_edition_years(f"{ev.get('title','')} {ev.get('rule','')}")
            eligible_years = {y for y in supported_years if y <= as_of_date.year}
            code_years = _extract_edition_years(code.get("name"))
            if eligible_years and code_years and not (code_years & eligible_years):
                target_year = max(eligible_years)
                edition_hint = any(
                    re.search(rf"\b{target_year}\b", f"{ev.get('title','')} {ev.get('rule','')}")
                    and re.search(r"\b(?:Oregon|state|county|city|adopted|adoption|code)\b", f"{ev.get('title','')} {ev.get('rule','')}", re.I)
                    for ev in currency_evidence
                )
                if edition_hint:
                    name = str(code.get("name") or "")
                    code["name"] = re.sub(r"\b(?:19|20)\d{2}\b", str(target_year), name, count=1)
                    code["validation_note"] = (
                        f"Code edition year corrected from the generated value using authoritative CODE_CURRENCY "
                        f"evidence supporting the {target_year} edition as of {as_of_date.isoformat()}."
                    )
        errs = code_currency_integrity_errors(code, evidence_by_id, as_of_date)
        if errs:
            code["status"] = "CONDITIONAL"
            code["validation_note"] = (
                "Downgraded from CURRENT because authoritative CODE_CURRENCY evidence did not "
                "establish that the stated edition is current as of the project date."
            )
    return data

def sanitize_unsupported_not_currently_triggered_statuses(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    changed = []
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown")
        app = item.get("applicability") or {}
        determination = app.get("determination")
        if determination == "cannot_determine":
            for key in ("permit", "pathway"):
                if item.get(key) == "NOT_CURRENTLY_TRIGGERED":
                    item[key] = "CONDITIONAL"
                    item[f"{key}_basis"] = "NOT_ESTABLISHED"
                    item[f"{key}_evidence"] = []
                    item[f"{key}_finding"] = (
                        f"The {discipline.lower()} {key} consequence is not established by current evidence "
                        "because the underlying applicability remains unresolved. Confirm the applicable "
                        "requirement using an authoritative, discipline-specific source."
                    )
                    changed.append(discipline)
            continue
        for key, required_type in (("permit", "PERMIT_EXEMPTION"), ("pathway", "PATHWAY")):
            if item.get(key) != "NOT_CURRENTLY_TRIGGERED":
                continue
            basis = item.get(f"{key}_basis")
            ids = item.get(f"{key}_evidence") or []
            valid = evidence_ids_supporting_type(ids, evidence_by_id, required_type, discipline)
            if determination != "applies" or basis == "NOT_ESTABLISHED" or not valid:
                item[key] = "CONDITIONAL"
                item[f"{key}_basis"] = "NOT_ESTABLISHED"
                item[f"{key}_evidence"] = []
                if key == "permit":
                    item[f"{key}_finding"] = (
                        f"A {discipline.lower()} permit is not currently established as exempt or "
                        "untriggered by current evidence. Confirm the permit requirement or exemption "
                        "using an authoritative permit-specific source."
                    )
                else:
                    item[f"{key}_finding"] = (
                        f"The {discipline.lower()} processing pathway is not currently established by "
                        "current evidence. Confirm the applicable pathway using an authoritative "
                        "discipline-specific source."
                    )
                changed.append(discipline)
    if changed:
        bottom = _norm_text(data.get("bottom_line"))
        if bottom:
            for discipline in set(changed):
                d = re.escape(discipline.lower())
                if re.search(rf"\b{d}\b[^.]*\b(?:not currently triggered|not currently required)\b", bottom, re.I):
                    data["bottom_line"] = (
                        "Current evidence establishes the governing research framework, but one or more "
                        "discipline-specific obligations remain conditional or not established. See the "
                        "Permit Matrix for the specific evidence and project facts needed to resolve them."
                    )
                    break
    return data

def sanitize_semantically_misplaced_permit_findings(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {e.get("id"): e for e in (data.get("evidence") or [])
                      if isinstance(e, dict) and e.get("id")}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        finding = _norm_text(item.get("permit_finding"))
        ids = item.get("permit_evidence") or []
        valid_permit = evidence_ids_supporting_type(ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline)
        valid_exemption = evidence_ids_supporting_type(ids, evidence_by_id, "PERMIT_EXEMPTION", discipline)
        if valid_permit or valid_exemption or not finding:
            continue
        if has_definitive_negative_permit_claim(finding):
            continue
        explicit_permit_consequence = any(re.search(p, finding) for p in [
            r"\bpermit\s+(?:is\s+)?required\b",
            r"\bpermit\s+requirement\s+(?:depends|turns)\s+on",
            r"\b(?:requires?|triggers?|necessitates?)\s+(?:a\s+)?(?:separate\s+)?permit\b",
            r"\b(?:a\s+)?permit\s+(?:would|will|may|might|could)\s+be\s+required\b",
            r"\bmust\s+obtain\s+(?:a\s+)?permit\b",
        ])
        epistemic = any(re.search(p, finding) for p in [
            r"\bpermit\s+(?:requirement\s+)?(?:is\s+)?not\s+established\b",
            r"\bcannot\s+determine\b[^.!?]{0,160}\bpermit\b",
            r"\bcurrent\s+evidence\b[^.!?]{0,180}\bdoes not establish\b[^.!?]{0,80}\bpermit\b",
        ])
        if explicit_permit_consequence or epistemic:
            continue
        review_or_compliance = any(re.search(p, finding) for p in [
            r"\breview\b", r"\bplan review\b", r"\bengineering\b",
            r"\bcompliance\b", r"\bcut sheets?\b", r"\bcomcheck\b",
            r"\binspection\b", r"\bcalculations?\b", r"\bsubmittal\b",
        ])
        if review_or_compliance:
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"The available evidence does not yet establish whether a {discipline.lower()} permit is required. "
                "Review or compliance obligations alone do not establish a permit requirement."
            )
    return data

def sanitize_unsubstantiated_conditional_statuses(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = item.get("type", "Unknown")
        permit_basis = item.get("permit_basis")
        permit_ids = item.get("permit_evidence") or []
        valid_permit = evidence_ids_supporting_type(
            permit_ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline
        ) or evidence_ids_supporting_type(
            permit_ids, evidence_by_id, "PERMIT_EXEMPTION", discipline
        )
        finding = _norm_text(item.get("permit_finding"))
        epistemic = any(re.search(p, finding) for p in [
            r"\bnot established\b", r"\bcannot determine\b",
            r"\bdoes not establish\b", r"\bnot determined\b",
        ])
        if item.get("permit") == "CONDITIONAL" and permit_basis == "NOT_ESTABLISHED" and not valid_permit and epistemic:
            item["permit"] = "UNKNOWN"
        pathway_basis = item.get("pathway_basis")
        pathway_ids = item.get("pathway_evidence") or []
        valid_pathway = evidence_ids_supporting_type(
            pathway_ids, evidence_by_id, "PATHWAY", discipline
        ) or evidence_ids_supporting_type(
            pathway_ids, evidence_by_id, "REVIEW_REQUIREMENT", discipline
        )
        pathway_finding = _norm_text(item.get("pathway_finding"))
        pathway_epistemic = any(re.search(p, pathway_finding) for p in [
            r"\bnot established\b", r"\bcannot determine\b",
            r"\bdoes not establish\b", r"\bunestablished\b",
        ])
        if item.get("pathway") == "CONDITIONAL" and pathway_basis == "NOT_ESTABLISHED" and not valid_pathway and pathway_epistemic:
            item["pathway"] = "UNKNOWN"
    return data

def sanitize_bottom_line_for_jurisdiction(data):
    if not isinstance(data, dict):
        return data
    jur = data.get("jurisdiction") or {}
    if str(jur.get("status", "")).upper() != "CONDITIONAL":
        return data
    text = str(data.get("bottom_line", ""))
    patterns = [
        r"\bis located in .*? under the jurisdiction of\b",
        r"\bis under the jurisdiction of\b",
        r"\bthe jurisdiction is\b",
    ]
    if any(re.search(p, text, flags=re.I) for p in patterns):
        county = jur.get("county") or "the identified county"
        ahj = jur.get("ahj") or "the identified AHJ"
        data["bottom_line"] = re.sub(
            r".*?(?=\bThe project is governed\b|$)",
            f"The address is treated as potentially within {county}, but jurisdiction remains conditional pending site-specific confirmation from {ahj}. ",
            text, count=1, flags=re.I | re.S
        )
    return data

def _discipline_permit_status_summary(data):
    summary = {"verified": [], "unresolved": [], "not_applicable": []}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("type") or "Unknown").strip()
        status = str(item.get("permit") or "UNKNOWN").upper()
        if status == "VERIFIED_REQUIRED":
            summary["verified"].append(name)
        elif status == "NOT_APPLICABLE":
            summary["not_applicable"].append(name)
        else:
            summary["unresolved"].append(name)
    return summary

def sanitize_bottom_line_against_final_matrix(data):
    if not isinstance(data, dict):
        return data
    summary = _discipline_permit_status_summary(data)
    bottom = _norm_text(data.get("bottom_line"))
    if not bottom or not summary["unresolved"]:
        return data
    permit_consequence = re.compile(
        r"\b(?:permit|permits|approval|approvals)\b[^.]{0,160}\b(?:required|requires|must|shall|verified|trigger(?:s|ed)?)\b|"
        r"\b(?:required|requires|must|shall|verified|trigger(?:s|ed)?)\b[^.]{0,160}\b(?:permit|permits|approval|approvals)\b",
        re.I,
    )
    if not permit_consequence.search(bottom):
        return data
    lowered = bottom.lower()
    names_in_bottom = [
        name for name in summary["unresolved"]
        if name and name.lower() in lowered
    ]
    blanket = bool(re.search(
        r"\b(?:the project|this project|the proposed .*?project)\b[^.]{0,100}\b(?:requires?|needs?|must obtain|has)\b[^.]{0,100}\b(?:permit|permits|approval|approvals)\b",
        lowered,
        re.I,
    ))
    if not names_in_bottom and not blanket:
        return data
    parts = []
    if summary["verified"]:
        parts.append("Permit requirements are established for " + ", ".join(summary["verified"]) + ".")
    if summary["unresolved"]:
        parts.append("Permit requirements remain unresolved for " + ", ".join(summary["unresolved"]) + ".")
    if summary["not_applicable"]:
        parts.append("No permit requirement is currently established for " + ", ".join(summary["not_applicable"]) + ".")
    parts.append("See the Permit Matrix and supporting evidence for the details and remaining research items.")
    data["bottom_line"] = " ".join(parts)
    return data



def _is_generic_regulatory_landing_page(evidence):
    """Return True when a regulatory URL is a generic agency/permit portal."""
    if not isinstance(evidence, dict):
        return False
    ptype = str(evidence.get("proposition_type") or "").upper()
    if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT"}:
        return False
    url = str(evidence.get("url") or "").strip()
    if not url:
        return False
    try:
        from urllib.parse import urlparse
        path = (urlparse(url).path or "/").rstrip("/").lower() or "/"
    except Exception:
        return False
    exact_generic = {
        "/", "/index.html", "/home", "/services", "/codes",
        "/building-and-safety", "/bsd", "/planning", "/building",
        "/permits", "/permit", "/permit-center", "/permitcenter",
        "/licensing", "/applications", "/application", "/applications-and-forms",
        "/applications-and-forms/", "/plumbing", "/mechanical", "/electrical", "/building",
        "/fire-prevention", "/fire", "/fire-safety", "/public-works",
        "/publicworks", "/engineering", "/planning-and-zoning",
        "/zoning", "/development-services", "/development",
        "/permits/epicla", "/permits/epicla/index.html",
    }
    if path in exact_generic:
        return True
    portal_suffixes = {"/epicla", "/permit-portal", "/permitportal", "/online-permits", "/online-permit", "/permit-system", "/permit-system-home"}
    return any(path.endswith(token) for token in portal_suffixes)

def sanitize_generic_proposition_links(data):
    """Strip generic landing pages from regulatory conclusion links without deleting evidence."""
    if not isinstance(data, dict):
        return data
    evidence_by_id = {e.get("id"): e for e in (data.get("evidence") or [])
                       if isinstance(e, dict) and e.get("id")}
    generic = {ptype: set() for ptype in ("PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT")}
    for eid, ev in evidence_by_id.items():
        if _is_generic_regulatory_landing_page(ev):
            generic[str(ev.get("proposition_type") or "").upper()].add(eid)
    if not any(generic.values()):
        return data

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown").strip()

        permit_ids = item.get("permit_evidence") or []
        if isinstance(permit_ids, str):
            permit_ids = [permit_ids]
        if isinstance(permit_ids, list):
            cleaned = [eid for eid in permit_ids
                       if eid not in generic["PERMIT_REQUIREMENT"]
                       and eid not in generic["PERMIT_EXEMPTION"]]
            item["permit_evidence"] = cleaned
            if str(item.get("permit") or "UNKNOWN").upper() == "VERIFIED_REQUIRED" and not evidence_ids_supporting_type(cleaned, evidence_by_id, "PERMIT_REQUIREMENT", discipline):
                item["permit"] = "CONDITIONAL"
                item["permit_basis"] = "NOT_ESTABLISHED"
                item["permit_evidence"] = []
                item["permit_finding"] = (
                    "The available evidence does not yet establish whether a permit is required. "
                    "Confirm the discipline-specific permit requirement using an authoritative permit-specific source."
                )

        pathway_ids = item.get("pathway_evidence") or []
        if isinstance(pathway_ids, str):
            pathway_ids = [pathway_ids]
        if isinstance(pathway_ids, list):
            cleaned_pathway = [eid for eid in pathway_ids if eid not in generic["PATHWAY"]]
            item["pathway_evidence"] = cleaned_pathway
            if str(item.get("pathway") or "UNKNOWN").upper() == "VERIFIED_REQUIRED" and not evidence_ids_supporting_type(cleaned_pathway, evidence_by_id, "PATHWAY", discipline):
                item["pathway"] = "CONDITIONAL"
                item["pathway_basis"] = "NOT_ESTABLISHED"
                item["pathway_evidence"] = []
                item["pathway_finding"] = (
                    "The available evidence does not yet establish the permitting pathway. "
                    "Confirm the applicable process using an authoritative pathway or review source."
                )

    bottom_ids = data.get("bottom_line_evidence") or []
    if isinstance(bottom_ids, str):
        bottom_ids = [bottom_ids]
    if isinstance(bottom_ids, list):
        data["bottom_line_evidence"] = [eid for eid in bottom_ids
            if eid not in generic["PERMIT_REQUIREMENT"]
            and eid not in generic["PERMIT_EXEMPTION"]
            and eid not in generic["PATHWAY"]
            and eid not in generic["ENTITLEMENT"]]
        if not data["bottom_line_evidence"]:
            fallback = next((eid for eid, ev in evidence_by_id.items() if not _is_generic_regulatory_landing_page(ev)), None)
            if fallback:
                data["bottom_line_evidence"] = [fallback]
    return data

def _grounding_web_sources(response):
    """Return official web sources exposed by Gemini grounding metadata.

    Gemini may place a vertexaisearch.cloud.google.com redirect URL in model JSON
    instead of the underlying source URL.  The grounding metadata normally carries
    the actual web URI; use it to repair the evidence record before validation.
    """
    sources = []
    try:
        candidate = (getattr(response, "candidates", None) or [None])[0]
        gm = getattr(candidate, "grounding_metadata", None) if candidate else None
        chunks = getattr(gm, "grounding_chunks", None) if gm else None
        for chunk in chunks or []:
            web = getattr(chunk, "web", None)
            if web is None and isinstance(chunk, dict):
                web = chunk.get("web")
            if web is None:
                continue
            uri = getattr(web, "uri", None) if not isinstance(web, dict) else web.get("uri")
            title = getattr(web, "title", None) if not isinstance(web, dict) else web.get("title")
            if uri and str(uri).startswith(("http://", "https://")):
                sources.append({"url": str(uri).strip(), "title": str(title or "").strip()})
    except Exception:
        return []
    return sources


def _is_grounding_redirect_url(url):
    value = str(url or "").strip().lower()
    return "vertexaisearch.cloud.google.com/grounding-api-redirect/" in value


def _repair_grounding_redirect_urls(data, response):
    """Replace Gemini grounding redirect URLs with the underlying grounded web URL.

    Match by source title first, then by distinctive words from the evidence title.
    Never invent a URL; if no grounded URL can be matched, leave the redirect in
    place so the deterministic source firewall can reject it.
    """
    if not isinstance(data, dict):
        return data
    sources = _grounding_web_sources(response)
    if not sources:
        return data
    evidence = data.get("evidence") or []
    for ev in evidence:
        if not isinstance(ev, dict) or not _is_grounding_redirect_url(ev.get("url")):
            continue
        title = _norm_text(ev.get("title"))
        if not title:
            continue
        best = None
        best_score = 0
        title_words = {w for w in re.findall(r"[a-z0-9]+", title) if len(w) >= 4}
        for src in sources:
            src_title = _norm_text(src.get("title"))
            if not src_title:
                continue
            src_words = {w for w in re.findall(r"[a-z0-9]+", src_title) if len(w) >= 4}
            score = len(title_words & src_words)
            if src_title == title:
                score += 100
            if score > best_score:
                best_score = score
                best = src
        if best and best_score >= (100 if _norm_text(best.get("title")) == title else 2):
            ev["url"] = best["url"]
    return data


def _normalize_evidence_urls(data):
    """Normalize common escaped URL forms without changing the target host/path."""
    if not isinstance(data, dict):
        return data
    for ev in data.get("evidence") or []:
        if not isinstance(ev, dict):
            continue
        url = str(ev.get("url") or "").strip()
        if not url:
            continue
        # Model JSON sometimes escapes punctuation for markdown/code presentation.
        url = url.replace("\\.", ".").replace("\\_", "_")
        ev["url"] = url
    return data


def evidence_source_specificity_errors(evidence):
    errors = []
    generic_paths = {
        "/", "/index.html", "/building-and-safety", "/bsd", "/codes",
    }
    for ev in evidence or []:
        if not isinstance(ev, dict):
            continue
        ptype = str(ev.get("proposition_type") or "").upper()
        if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT"}:
            continue
        url = str(ev.get("url") or "").strip()
        if _is_grounding_redirect_url(url):
            errors.append(f"{ev.get('id', 'Evidence')}: {ptype} evidence still points to a Gemini grounding redirect, not the underlying source URL.")
            continue
        if not url:
            errors.append(f"{ev.get('id', 'Evidence')}: {ptype} evidence has no URL.")
            continue
        try:
            from urllib.parse import urlparse
            parsed = urlparse(url)
            path = (parsed.path or "/").rstrip("/") or "/"
            host = (parsed.netloc or "").lower()
        except Exception:
            continue
        if path in generic_paths or _is_generic_regulatory_landing_page(ev):
            errors.append(f"{ev.get('id', 'Evidence')}: {ptype} evidence points to a generic agency/permit-portal page ({host}{path}), not a proposition-specific source.")
            continue

        # A non-generic URL is not sufficient by itself.  The extracted rule
        # must contain the proposition being claimed.  This prevents a specific
        # agency page that merely discusses code compliance/review from being
        # promoted into permit or entitlement evidence.  The detailed lexical
        # checks remain in evidence_proposition_integrity_errors(); this check
        # makes source specificity enforce the same semantic contract.
        rule = _norm_text(ev.get("rule"))

        # A discipline landing page such as /plumbing/ or /fire-prevention/
        # cannot be promoted to proposition-specific evidence merely because
        # Gemini gave it a permit-related title.  A specific permit/checklist/
        # application/rule document path is required for permit propositions.
        if ptype in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION"}:
            url_signal = re.search(
                r"(?:permit|application|requirements?|checklist|plan[-_ ]?review|fire[-_ ]?alarm|plumbing[-_ ]?permit|mechanical[-_ ]?permit|electrical[-_ ]?permit)",
                path, re.I,
            )
            document_signal = re.search(r"\.(?:pdf|docx?|xlsx?)$", path, re.I)
            if not url_signal and not document_signal:
                errors.append(
                    f"{ev.get('id', 'Evidence')}: {ptype} URL does not identify a permit-specific page/document; generic discipline pages are insufficient."
                )
                continue

        if ptype == "PERMIT_REQUIREMENT" and not re.search(
            r"\b(?:permit|approval|license)\b[^.!?;:]{0,220}\b(?:required|needed|necessary)\b|"
            r"\b(?:requires?|must obtain|shall obtain)\b[^.!?;:]{0,220}\b(?:permit|approval|license)\b",
            rule, re.I
        ):
            errors.append(f"{ev.get('id', 'Evidence')}: PERMIT_REQUIREMENT source rule does not itself state a permit/approval/license requirement.")
        elif ptype == "PERMIT_EXEMPTION" and not re.search(
            r"\b(?:exempt(?:ed|ion)?|no\s+(?:separate\s+)?permit|permit\s+(?:is\s+)?not\s+required|does\s+not\s+require\s+(?:a\s+)?permit|not\s+subject\s+to\s+(?:a\s+)?permit)\b",
            rule, re.I
        ):
            errors.append(f"{ev.get('id', 'Evidence')}: PERMIT_EXEMPTION source rule does not itself state an explicit exemption/non-permit proposition.")
        elif ptype == "PATHWAY" and not re.search(
            r"\b(?:submit|file|portal|permit center|online|application|plan review|processed|processing|inspection)\b",
            rule, re.I
        ):
            errors.append(f"{ev.get('id', 'Evidence')}: PATHWAY source rule does not itself state a processing/submittal/review pathway.")
    return errors

def sanitize_bottom_line_for_unestablished_permits(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    bottom = _norm_text(data.get("bottom_line"))
    if not bottom:
        return data
    unresolved = []
    verified = []
    not_applicable = []
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown").strip()
        status = str(item.get("permit") or "UNKNOWN").upper()
        if status == "VERIFIED_REQUIRED":
            verified.append(discipline)
        elif status == "NOT_APPLICABLE":
            not_applicable.append(discipline)
        else:
            unresolved.append(discipline)
    definitive_markers = re.compile(
        r"\b(?:permit|permits|approval|approvals)\b[^.]{0,100}\b"
        r"(?:is|are|must|shall|requires?|required|verified|trigger(?:s|ed)?)\b",
        re.I,
    )
    if unresolved and definitive_markers.search(bottom):
        parts = []
        if verified:
            parts.append("Proposition-specific permit evidence was found for " + ", ".join(verified) + ".")
        parts.append("Other permit questions remain open where the available evidence does not establish the legal consequence.")
        parts.append("Use the discipline questions, expected AHJ process, and clearly labeled possible leads to continue the research.")
        data["bottom_line"] = " ".join(parts)
    return data

def sanitize_unverifiable_verified_permits(data):
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    changed_disciplines = []
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown")
        if item.get("permit") != "VERIFIED_REQUIRED":
            continue
        ids = item.get("permit_evidence") or []
        valid = evidence_ids_supporting_type(
            ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline
        )
        if valid:
            continue
        item["permit"] = "CONDITIONAL"
        item["permit_basis"] = "NOT_ESTABLISHED"
        item["permit_evidence"] = []
        item["permit_finding"] = (
            f"The {discipline.lower()} permit requirement is not established by "
            "current permit-specific evidence and remains conditional pending "
            "confirmation from an authoritative permit-specific source."
        )
        changed_disciplines.append(discipline)
    if not changed_disciplines:
        return data
    bottom_line = _norm_text(data.get("bottom_line"))
    if bottom_line:
        lowered = bottom_line.lower()
        needs_rewrite = False
        for discipline in changed_disciplines:
            d = discipline.lower()
            patterns = [
                rf"\b{re.escape(d)}\b[^.]*\b(?:permit|approval)\b[^.]*\b(?:requires?|required|must obtain|triggers?)\b",
                rf"\b(?:permit|approval)\b[^.]*\b(?:requires?|required|must obtain|triggers?)\b[^.]*\b{re.escape(d)}\b",
            ]
            if any(re.search(p, lowered, re.I) for p in patterns):
                needs_rewrite = True
                break
        if needs_rewrite:
            data["bottom_line"] = (
                "Current evidence establishes the governing research framework, "
                "but one or more discipline-specific permit requirements remain "
                "conditional or not established. See the Permit Matrix for the "
                "specific missing facts and evidence needed to resolve them."
            )
            existing = [
                eid for eid in (data.get("bottom_line_evidence") or [])
                if eid in evidence_by_id
            ]
            if existing:
                data["bottom_line_evidence"] = existing
            elif evidence_by_id:
                data["bottom_line_evidence"] = [next(iter(evidence_by_id))]
    return data


def sanitize_validation_blocking_permit_findings(data):
    """Final deterministic permit contract firewall.

    This runs after all Gemini output/repair normalization.  The validator's
    contract is simple: a permit conclusion can only be definitive when the
    discipline has valid permit-specific evidence (or an explicit exemption).
    Do not ask Gemini to repair this deterministic condition.  If a discipline
    lacks that evidence, neutralize any non-NA permit conclusion/finding here.
    """
    if not isinstance(data, dict):
        return data

    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    changed = []
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue

        discipline = str(item.get("type") or "Unknown").strip()
        permit = str(item.get("permit") or "UNKNOWN").strip().upper()
        ids = item.get("permit_evidence") or []
        if isinstance(ids, str):
            ids = [ids]
        if not isinstance(ids, list):
            ids = []

        valid_permit = evidence_ids_supporting_type(
            ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline
        )
        valid_exemption = evidence_ids_supporting_type(
            ids, evidence_by_id, "PERMIT_EXEMPTION", discipline
        )

        # NOT_APPLICABLE is governed by the explicit exemption validator.
        # UNKNOWN is already epistemically neutral.  Everything else needs
        # actual permit-specific support before it can make a permit claim.
        if permit in {"NOT_APPLICABLE", "UNKNOWN"}:
            continue
        if valid_permit or valid_exemption:
            continue

        item["permit"] = "CONDITIONAL"
        item["permit_basis"] = "NOT_ESTABLISHED"
        item["permit_evidence"] = []
        item["permit_finding"] = (
            f"The available evidence does not yet establish whether a "
            f"{discipline.lower()} permit is required. Confirm the applicable "
            "discipline-specific permit requirement using an authoritative "
            "permit-specific source."
        )
        changed.append(discipline)

    if changed:
        # A permit conclusion may have been rewritten after the earlier
        # bottom-line sanitizers ran. Reconcile it once more, deterministically.
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)

    return data

def sanitize_final_regulatory_statuses(data):
    """Last deterministic firewall before validation.

    Gemini repair can occasionally put an evidence-basis token such as
    DIRECT_EVIDENCE into the permit/pathway status field, or reintroduce a
    permit consequence after earlier sanitizers ran.  Never let those model
    formatting errors reach the validator or trigger another paid repair.
    """
    if not isinstance(data, dict):
        return data

    allowed_statuses = {
        "VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED", "UNKNOWN",
        "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED",
    }
    evidence_by_id = {
        e.get("id"): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    permit_claim_patterns = [
        r"\bpermit\s+(?:is\s+)?required\b",
        r"\bpermit\s+(?:is\s+)?required\s+(?:if|when|once|where)\b",
        r"\bpermit\s+requirement\s+(?:depends|turns)\s+on",
        r"\b(?:requires?|triggers?|necessitates?)\s+(?:a\s+)?(?:separate\s+)?permit\b",
        r"\b(?:a\s+)?permit\s+(?:would|will|may|might|could)\s+be\s+required\b",
        r"\bmust\s+obtain\s+(?:a\s+)?permit\b",
        r"\b(?:permit|approval)\s+is\s+triggered\s+by\b",
    ]
    pathway_claim_patterns = [
        r"\b(?:submit|file)\b",
        r"\b(?:through|via|using)\s+(?:the\s+)?(?:portal|permit\s+center|online)\b",
        r"\bprocessed\s+(?:as|through|via)\b",
        r"\bplan\s+review\b",
        r"\binspection\s+required\b",
        r"\bpathway\b",
    ]

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue

        discipline = str(item.get("type") or "Unknown")
        family = discipline.lower()

        # Repair models sometimes copy permit_basis into permit itself.
        permit = str(item.get("permit") or "").strip().upper()
        if permit not in allowed_statuses:
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"The {family} permit requirement is not established by the "
                "available permit-specific evidence. Confirm the applicable "
                "requirement using an authoritative permit-specific source."
            )
            permit = "CONDITIONAL"

        pathway = str(item.get("pathway") or "").strip().upper()
        if pathway not in allowed_statuses:
            item["pathway"] = "CONDITIONAL"
            item["pathway_basis"] = "NOT_ESTABLISHED"
            item["pathway_evidence"] = []
            item["pathway_finding"] = (
                f"The {family} processing pathway is not established by the "
                "available pathway-specific evidence. Confirm the applicable "
                "AHJ process using an authoritative source."
            )
            pathway = "CONDITIONAL"

        permit_ids = item.get("permit_evidence") or []
        if isinstance(permit_ids, str):
            permit_ids = [permit_ids]
        pathway_ids = item.get("pathway_evidence") or []
        if isinstance(pathway_ids, str):
            pathway_ids = [pathway_ids]

        valid_permit = evidence_ids_supporting_type(
            permit_ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline
        )
        valid_exemption = evidence_ids_supporting_type(
            permit_ids, evidence_by_id, "PERMIT_EXEMPTION", discipline
        )
        valid_pathway = evidence_ids_supporting_type(
            pathway_ids, evidence_by_id, "PATHWAY", discipline
        )
        valid_review = evidence_ids_supporting_type(
            pathway_ids, evidence_by_id, "REVIEW_REQUIREMENT", discipline
        )

        permit_finding = _norm_text(item.get("permit_finding"))
        epistemic = any(re.search(p, permit_finding, re.I) for p in [
            r"\bnot established\b",
            r"\bcannot determine\b",
            r"\bunable to determine\b",
            r"\bdoes not establish\b",
            r"\bwhether\b[^.]{0,120}\bpermit\s+is\s+required\b",
        ])
        claims_permit = any(
            re.search(p, permit_finding, re.I) for p in permit_claim_patterns
        )

        # A permit conclusion without valid permit-specific evidence is never
        # allowed to remain verified/inferred or direct-evidence-backed.
        if (
            not valid_permit
            and not valid_exemption
            and permit != "NOT_APPLICABLE"
            and (
                permit in {"VERIFIED_REQUIRED", "INFERRED"}
                or str(item.get("permit_basis") or "").upper() == "DIRECT_EVIDENCE"
                or (claims_permit and not epistemic)
            )
        ):
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"The available evidence does not yet establish whether a "
                f"{family} permit is required. Check the applicable "
                "discipline-specific permit requirement against an authoritative source."
            )

        pathway_finding = _norm_text(item.get("pathway_finding"))
        claims_pathway = any(
            re.search(p, pathway_finding, re.I) for p in pathway_claim_patterns
        )
        if (
            not valid_pathway
            and not valid_review
            and pathway != "NOT_APPLICABLE"
            and (
                pathway in {"VERIFIED_REQUIRED", "INFERRED"}
                or str(item.get("pathway_basis") or "").upper() == "DIRECT_EVIDENCE"
                or (claims_pathway and not re.search(
                    r"\b(?:not established|cannot determine|unable to determine)\b",
                    pathway_finding, re.I
                ))
            )
        ):
            item["pathway"] = "CONDITIONAL"
            item["pathway_basis"] = "NOT_ESTABLISHED"
            item["pathway_evidence"] = []
            item["pathway_finding"] = (
                f"The {family} processing pathway is not established by current "
                "evidence. Confirm the applicable AHJ process using an authoritative source."
            )

    return data

def sanitize_contradictory_established_pathways(data):
    """Downgrade pathways whose own applicability text admits no specific rule/evidence.

    This is a deterministic consistency repair, not a research judgment: an
    established/verified pathway must be backed by an actual pathway proposition.
    """
    if not isinstance(data, dict):
        return data
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        if str(item.get("pathway") or "").upper() != "VERIFIED_REQUIRED":
            continue
        applicability = item.get("applicability") or {}
        if not isinstance(applicability, dict):
            continue
        text = _norm_text(" ".join([
            str(applicability.get("rule") or ""),
            str(applicability.get("rule_finding") or ""),
            str(applicability.get("source_finding") or ""),
        ]))
        if re.search(
            r"\bno\s+proposition[- ]specific\s+(?:authoritative\s+)?(?:rule|evidence)\s+was\s+established\b",
            text, re.I
        ):
            discipline = str(item.get("type") or "Unknown").strip().lower()
            item["pathway"] = "CONDITIONAL"
            item["pathway_basis"] = "NOT_ESTABLISHED"
            item["pathway_evidence"] = []
            item["pathway_finding"] = (
                f"The {discipline} processing pathway is not established by current evidence. "
                "Confirm the applicable submission or review process with an authoritative source."
            )
    return data

def validate_dossier(data):
    errors = []
    allowed_statuses = {"VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED", "UNKNOWN", "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED"}
    allowed_determinations = {"applies", "does_not_apply", "cannot_determine"}
    allowed_relationships = {"direct", "conditional", "not_established"}
    allowed_completeness = {"SUFFICIENT", "PARTIAL", "INSUFFICIENT"}
    allowed_evidence_basis = {"DIRECT_EVIDENCE", "CONDITIONAL", "NOT_ESTABLISHED"}

    completeness = data.get("research_completeness", {})
    if completeness.get("status") not in allowed_completeness:
        errors.append("Research completeness must be SUFFICIENT, PARTIAL, or INSUFFICIENT.")
    if completeness.get("status") == "INSUFFICIENT" and not completeness.get("reason"):
        errors.append("INSUFFICIENT research completeness requires a reason.")
    evidence_items = data.get("evidence", [])
    evidence_by_id = {
        e.get("id"): e
        for e in evidence_items
        if e.get("id")
    }
    evidence_ids = set(evidence_by_id.keys())

    # Source specificity is enforced when evidence is actually used to support
    # a regulatory conclusion. An unused generic agency page is not itself a
    # regulatory inconsistency.
    referenced_proposition_ids = set()

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue

        for field in (
            "permit_evidence",
            "pathway_evidence",
            "authority_evidence",
        ):
            raw_ids = item.get(field) or []

            if isinstance(raw_ids, str):
                raw_ids = [raw_ids]

            if isinstance(raw_ids, list):
                referenced_proposition_ids.update(
                    eid
                    for eid in raw_ids
                    if isinstance(eid, str)
                )

        applicability = item.get("applicability") or {}

        if isinstance(applicability, dict):
            raw_ids = applicability.get("evidence") or []

            if isinstance(raw_ids, str):
                raw_ids = [raw_ids]

            if isinstance(raw_ids, list):
                referenced_proposition_ids.update(
                    eid
                    for eid in raw_ids
                    if isinstance(eid, str)
                )

    for eid in data.get("bottom_line_evidence") or []:
        if isinstance(eid, str):
            referenced_proposition_ids.add(eid)

    for evidence in evidence_items:
        if evidence.get("id") not in referenced_proposition_ids:
            continue

        errors.extend(
            evidence_source_specificity_errors([evidence])
        )

    for ev in evidence_items:
        eid = ev.get("id", "Unknown")
        if not ev.get("title"):
            errors.append(f"{eid}: missing title.")
        if not ev.get("url"):
            errors.append(f"{eid}: missing URL.")
        if not ev.get("authority"):
            errors.append(f"{eid}: missing authority.")
        if not ev.get("discipline"):
            errors.append(f"{eid}: missing discipline.")
        if not ev.get("rule"):
            errors.append(f"{eid}: missing rule proposition.")
        proposition_type = ev.get("proposition_type")
        if proposition_type not in EVIDENCE_PROPOSITION_TYPES:
            errors.append(
                f"{eid}: invalid or missing proposition_type "
                f"'{proposition_type}'."
            )
        errors.extend(evidence_proposition_integrity_errors(ev))
        errors.extend(jurisdiction_evidence_integrity_errors(ev))

    jurisdiction = data.get("jurisdiction", {})
    for eid in jurisdiction.get("evidence", []):
        if eid not in evidence_ids: errors.append(f"Jurisdiction references nonexistent evidence: {eid}")
    if jurisdiction.get("status") == "VERIFIED" and not jurisdiction.get("evidence"):
        errors.append("Verified jurisdiction must have evidence.")
    if jurisdiction.get("status") == "VERIFIED":
        valid_jurisdiction_ids = [
            eid for eid in (jurisdiction.get("evidence") or [])
            if eid in evidence_by_id
            and evidence_by_id[eid].get("proposition_type") == "JURISDICTION"
            and not evidence_proposition_integrity_errors(evidence_by_id[eid])
            and not jurisdiction_evidence_integrity_errors(evidence_by_id[eid])
        ]
        if not valid_jurisdiction_ids:
            errors.append("Verified jurisdiction requires valid JURISDICTION evidence.")
        if not any(_jurisdiction_evidence_is_site_specific(evidence_by_id[eid], str(data.get("address") or "")) for eid in valid_jurisdiction_ids):
            errors.append("Verified jurisdiction requires site-specific parcel/boundary evidence; postal city or generic agency coverage is insufficient.")

    for code in data.get("codes", []):
        for eid in code.get("evidence", []):
            if eid not in evidence_ids: errors.append(f"Code '{code.get('name')}' references nonexistent evidence: {eid}")
        if code.get("status") == "CURRENT" and not code.get("evidence"):
            errors.append(f"Current code '{code.get('name')}' must have evidence.")
        errors.extend(code_currency_integrity_errors(code, evidence_by_id))

    bottom_line_evidence = data.get("bottom_line_evidence") or []
    if not bottom_line_evidence:
        errors.append("Bottom Line: must include bottom_line_evidence.")
    else:
        for evidence_id in bottom_line_evidence:
            if evidence_id not in evidence_by_id:
                errors.append(f"Bottom Line: unknown evidence ID '{evidence_id}'.")
                
    bottom_line = (data.get("bottom_line") or "").lower()
    material_markers = ["permit", "required", "requires", "code", "ahj", "jurisdiction", "conditional use", "cup", "review", "approval"]
    if any(marker in bottom_line for marker in material_markers):
        if not bottom_line_evidence:
            errors.append("Bottom Line: material regulatory conclusions require evidence.")

    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        permit_basis = item.get("permit_basis")
        pathway_basis = item.get("pathway_basis")
        permit_finding = item.get("permit_finding", "")
        pathway_finding = item.get("pathway_finding", "")
        permit_evidence_ids = item.get("permit_evidence") or []
        pathway_evidence_ids = item.get("pathway_evidence") or []
        app = item.get("applicability", {})
        determination = app.get("determination")
        relationship = app.get("relationship")
        fact = app.get("fact", {})
        fact_source = fact.get("source") if isinstance(fact, dict) else None
        fact_statement = fact.get("statement", "") if isinstance(fact, dict) else str(fact)

        if permit not in allowed_statuses: errors.append(f"{discipline}: invalid permit '{permit}'")
        if pathway not in allowed_statuses: errors.append(f"{discipline}: invalid pathway '{pathway}'")
        if determination not in allowed_determinations: errors.append(f"{discipline}: invalid determination '{determination}'")
        if relationship not in allowed_relationships: errors.append(f"{discipline}: invalid relationship '{relationship}'")
        if permit_basis not in allowed_evidence_basis: errors.append(f"{discipline}: invalid permit_basis '{permit_basis}'.")
        if pathway_basis not in allowed_evidence_basis: errors.append(f"{discipline}: invalid pathway_basis '{pathway_basis}'.")
        if fact_source not in FACT_SOURCES: errors.append(f"{discipline}: invalid fact source '{fact_source}'")
        if not fact_statement: errors.append(f"{discipline}: applicability fact is missing.")

        if fact_source in {"INFERRED", "UNKNOWN"}:
            if determination in {"applies", "does_not_apply"} and relationship == "direct":
                errors.append(f"{discipline}: {fact_source.lower()} fact cannot establish direct applicability.")
            if determination == "does_not_apply":
                errors.append(f"{discipline}: {fact_source.lower()} fact cannot establish non-applicability.")

        for eid in app.get("evidence", []):
            if eid not in evidence_ids: errors.append(f"{discipline}: applicability references nonexistent evidence '{eid}'")
            evidence = evidence_by_id.get(eid)
            if evidence and relationship == "direct":
                ev_disc = (evidence.get("discipline") or "").strip().lower()
                current_disc = discipline.strip().lower()
                ev_family = discipline_family(ev_disc)
                current_family = discipline_family(current_disc)
                
                # 3. Validator softened: do not append an error, just continue
                if ev_family and current_family and ev_family != current_family:
                    continue

        for eid in permit_evidence_ids:
            if eid not in evidence_ids: errors.append(f"{discipline}: permit references nonexistent evidence '{eid}'")
        for eid in pathway_evidence_ids:
            if eid not in evidence_ids: errors.append(f"{discipline}: pathway references nonexistent evidence '{eid}'")

        authority_evidence_ids = item.get("authority_evidence") or []
        if isinstance(authority_evidence_ids, str):
            authority_evidence_ids = [authority_evidence_ids]

        if permit == "VERIFIED_REQUIRED":
            if permit_basis != "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: VERIFIED_REQUIRED permit must have permit_basis=DIRECT_EVIDENCE.")
            if not permit_evidence_ids:
                errors.append(f"{discipline}: VERIFIED_REQUIRED permit requires permit-specific evidence.")
            else:
                supported = False
                for eid in permit_evidence_ids:
                    evidence = evidence_by_id.get(eid)
                    if evidence_supports_permit_requirement(evidence, discipline):
                        supported = True
                        break
                if not supported:
                    errors.append(f"{discipline}: VERIFIED_REQUIRED permit claim is not supported by PERMIT_REQUIREMENT evidence.")
            
            state_permit_ids = [
                eid for eid in permit_evidence_ids
                if eid in evidence_by_id
                and str(evidence_by_id[eid].get("authority") or "").lower() == "state"
                and evidence_supports_permit_requirement(evidence_by_id[eid], discipline)
            ]
            if state_permit_ids:
                authority_evidence_ids = item.get("authority_evidence") or []
                if isinstance(authority_evidence_ids, str):
                    authority_evidence_ids = [authority_evidence_ids]
                if not any(evidence_supports_authority_hierarchy(evidence_by_id.get(eid)) for eid in authority_evidence_ids):
                    item.setdefault("validation_notes", []).append(
                        "No separate AUTHORITY_HIERARCHY evidence was linked. Verify local amendments, "
                        "delegated authority, or home-rule/local-override provisions where applicable."
                    )

            permit_text = str(permit_finding or "").lower()
            permit_finding_is_epistemic = is_epistemic_permit_finding(permit_text)
            conditional_trigger_patterns = [
                r"\bpermit\s+(?:is\s+)?required\s+if\b",
                r"\bpermit\s+(?:is\s+)?required\s+when\b",
                r"\bpermit\s+(?:is\s+)?required\s+only\s+if\b",
                r"\btriggers?\s+(?:a\s+)?permit\b",
                r"\brequires?\s+(?:a\s+)?permit\b",
            ]
            definitive_permit_consequence = (
                not permit_finding_is_epistemic
                and any(re.search(pattern, permit_text) for pattern in conditional_trigger_patterns)
            )
            if definitive_permit_consequence:
                if not evidence_supports_any(permit_evidence_ids, evidence_by_id, {"PERMIT_REQUIREMENT"}, discipline):
                    errors.append(f"{discipline}: permit finding asserts a permit trigger but lacks PERMIT_REQUIREMENT evidence.")
            
            exemption_patterns = [
                r"\bexempt(?:ed|ion)?\b", r"\bno\s+(?:separate\s+)?permit\b",
                r"\bpermit\s+is\s+not\s+required\b", r"\bdoes\s+not\s+require\s+(?:a\s+)?permit\b",
                r"\bnot\s+subject\s+to\s+(?:a\s+)?permit\b",
            ]
            has_exemption_claim = any(re.search(pattern, permit_text) for pattern in exemption_patterns)
            if has_exemption_claim and permit not in {"UNKNOWN", "CONDITIONAL"}:
                if not evidence_supports_any(permit_evidence_ids, evidence_by_id, {"PERMIT_EXEMPTION"}, discipline):
                    errors.append(f"{discipline}: definitive exemption/non-permit claim lacks PERMIT_EXEMPTION evidence.")

            if pathway_finding_text := str(pathway_finding or "").lower():
                pathway_claims_permit = re.search(
                    r"\bpermit\s+(?:is\s+)?required\b|\brequires?\s+(?:a\s+)?permit\b",
                    pathway_finding_text
                )
                if (
                    permit not in {"VERIFIED_REQUIRED", "NOT_APPLICABLE"}
                    and pathway_claims_permit
                    and not is_epistemic_permit_finding(pathway_finding_text)
                ):
                    errors.append(f"{discipline}: pathway_finding contains a permit requirement claim while the permit is not established.")

        if pathway == "VERIFIED_REQUIRED":
            if pathway_basis != "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: VERIFIED_REQUIRED pathway must have pathway_basis=DIRECT_EVIDENCE.")
            if not pathway_evidence_ids:
                errors.append(f"{discipline}: VERIFIED_REQUIRED pathway requires pathway-specific evidence.")
            else:
                supported = False
                for eid in pathway_evidence_ids:
                    evidence = evidence_by_id.get(eid)
                    if evidence_supports_pathway(evidence, discipline):
                        supported = True
                        break
                if not supported:
                    errors.append(f"{discipline}: VERIFIED_REQUIRED pathway claim is not supported by PATHWAY evidence.")

        if permit_basis == "DIRECT_EVIDENCE" and not permit_evidence_ids:
            errors.append(f"{discipline}: DIRECT_EVIDENCE permit basis requires permit_evidence.")
        if pathway_basis == "DIRECT_EVIDENCE" and not pathway_evidence_ids:
            errors.append(f"{discipline}: DIRECT_EVIDENCE pathway basis requires pathway_evidence.")

        if relationship == "direct" and not app.get("rule"): errors.append(f"{discipline}: direct relationship requires a source rule.")
        if determination == "cannot_determine" and not app.get("missing"): errors.append(f"{discipline}: cannot_determine requires missing facts.")
        if determination == "does_not_apply" and app.get("missing"):
            errors.append(f"{discipline}: determination=does_not_apply is invalid when applicability depends on unresolved facts.")
        if determination == "cannot_determine" and permit == "NOT_APPLICABLE":
            errors.append(f"{discipline}: permit=NOT_APPLICABLE is invalid when applicability cannot be determined.")
        if determination == "cannot_determine" and pathway == "NOT_APPLICABLE":
            errors.append(f"{discipline}: pathway=NOT_APPLICABLE is invalid when applicability cannot be determined.")
        
        if permit == "NOT_APPLICABLE":
            if determination != "does_not_apply": errors.append(f"{discipline}: NOT_APPLICABLE permit status requires determination=does_not_apply.")
            if relationship != "direct": errors.append(f"{discipline}: NOT_APPLICABLE permit status requires relationship=direct.")
            if not app.get("evidence"): errors.append(f"{discipline}: NOT_APPLICABLE permit status requires authoritative applicability evidence.")
            if app.get("missing"): errors.append(f"{discipline}: NOT_APPLICABLE permit status cannot have unresolved applicability facts.")

        if pathway == "NOT_APPLICABLE":
            if determination != "does_not_apply": errors.append(f"{discipline}: NOT_APPLICABLE pathway status requires determination=does_not_apply.")
            if relationship != "direct": errors.append(f"{discipline}: NOT_APPLICABLE pathway status requires relationship=direct.")
            if not app.get("evidence"): errors.append(f"{discipline}: NOT_APPLICABLE pathway status requires authoritative applicability evidence.")
            if app.get("missing"): errors.append(f"{discipline}: NOT_APPLICABLE pathway status cannot have unresolved applicability facts.")

        if permit == "NOT_CURRENTLY_TRIGGERED":
            if determination == "does_not_apply" and app.get("missing"):
                errors.append(f"{discipline}: NOT_CURRENTLY_TRIGGERED cannot accompany does_not_apply when applicability facts remain unresolved.")
            if determination == "cannot_determine":
                errors.append(f"{discipline}: NOT_CURRENTLY_TRIGGERED should not be used when the underlying applicability determination is still unknown.")

        if relationship == "not_established":
            finding = (permit_finding or "").lower()
            # Epistemic wording can contain the literal substring "is required"
            # (e.g. "whether a permit is required") without asserting a permit
            # consequence. Never flag that wording as a downstream claim.
            if not is_epistemic_permit_finding(finding):
                for phrase in ["is required", "requires", "shall require", "automatically triggers", "therefore requires", "must obtain"]:
                    if phrase in finding:
                        errors.append(f"{discipline}: relationship is not_established but permit_finding claims downstream requirement.")
                        break

        permit_finding_text = str(permit_finding or "").strip().lower()
        pathway_finding_text = str(pathway_finding or "").strip().lower()
        negative_pathway_patterns = [
            r"\bno plan review is required\b", r"\bno plan review required\b", r"\bplan review is not required\b",
            r"\bplan review not required\b", r"\bdoes not require plan review\b", r"\bdoes not trigger plan review\b",
            r"\bno land use approval is required\b", r"\bland use approval is not required\b",
            r"\bno cup amendment is required\b", r"\bcup amendment is not required\b",
            r"\bno cup modification is required\b", r"\bcup modification is not required\b",
        ]
        has_definitive_negative_permit = has_definitive_negative_permit_claim(permit_finding_text)
        has_definitive_negative_pathway = any(re.search(pattern, pathway_finding_text) for pattern in negative_pathway_patterns)

        if has_definitive_negative_permit:
            negative_supported = any(
                evidence_supports_permit_exemption(evidence_by_id.get(eid), discipline)
                for eid in permit_evidence_ids
            )
            if not negative_supported:
                errors.append(f"{discipline}: definitive negative permit conclusion requires PERMIT_EXEMPTION evidence.")
        if has_definitive_negative_pathway:
            negative_pathway_supported = any(
                evidence_by_id.get(eid, {}).get("proposition_type") == "PATHWAY"
                for eid in pathway_evidence_ids
            )
            if not negative_pathway_supported:
                errors.append(f"{discipline}: definitive negative pathway conclusion requires PATHWAY evidence.")

        regulatory_text = " ".join([str(permit_finding or ""), str(pathway_finding or "")])
        concrete_threshold_claim = has_concrete_threshold_claim(regulatory_text, discipline)
        if concrete_threshold_claim:
            threshold_evidence = any(
                evidence_supports_threshold(evidence_by_id.get(eid), discipline)
                for eid in (permit_evidence_ids + pathway_evidence_ids + (app.get("evidence") or []))
            )
            if not threshold_evidence:
                errors.append(f"{discipline}: concrete threshold/limit claim lacks discipline-matched THRESHOLD evidence.")

        review_markers = ["plan review", "engineering review", "inspection required", "administrative review", "review required"]
        if any(marker in pathway_finding_text for marker in review_markers):
            has_review_or_pathway = evidence_supports_any(pathway_evidence_ids, evidence_by_id, {"REVIEW_REQUIREMENT", "PATHWAY"}, discipline)
            if not has_review_or_pathway and pathway_basis == "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: pathway finding asserts a review/process requirement without REVIEW_REQUIREMENT or PATHWAY evidence.")

        if pathway == "VERIFIED_REQUIRED" and relationship == "not_established":
            errors.append(f"{discipline}: pathway cannot be verified when relationship is not_established.")

        if permit == "NOT_APPLICABLE":
            if not evidence_supports_any(permit_evidence_ids, evidence_by_id, {"PERMIT_EXEMPTION"}, discipline):
                errors.append(f"{discipline}: NOT_APPLICABLE permit requires explicit PERMIT_EXEMPTION evidence.")
            if permit_basis != "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: NOT_APPLICABLE permit must use permit_basis=DIRECT_EVIDENCE.")

        if pathway == "NOT_APPLICABLE":
            if not evidence_supports_any(pathway_evidence_ids, evidence_by_id, {"PATHWAY"}, discipline):
                errors.append(f"{discipline}: NOT_APPLICABLE pathway requires explicit PATHWAY evidence establishing exclusion/non-applicability.")

    for item in data.get("disciplines", []):
        errors.extend(semantic_consequence_errors(item, evidence_by_id))

    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        permit_finding = _norm_text(item.get("permit_finding"))
        pathway_finding = _norm_text(item.get("pathway_finding"))
        permit_ids = item.get("permit_evidence") or []
        pathway_ids = item.get("pathway_evidence") or []
        valid_permit_ids = evidence_ids_supporting_type(permit_ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline)
        valid_exemption_ids = evidence_ids_supporting_type(permit_ids, evidence_by_id, "PERMIT_EXEMPTION", discipline)
        valid_pathway_ids = evidence_ids_supporting_type(pathway_ids, evidence_by_id, "PATHWAY", discipline)
        trigger_patterns = [
            r"\bpermit\s+(?:is\s+)?required\b",
            r"\bpermit\s+required\b",
            r"\brequires?\s+(?:a\s+)?permit\b",
            r"\btriggers?\s+(?:a\s+)?permit\b",
            r"\bmust\s+obtain\s+(?:a\s+)?permit\b",
        ]
        finding_claims_permit = any(re.search(p, permit_finding) for p in trigger_patterns)
        # Do not treat epistemic language such as "does not yet establish whether
        # a permit is required" as a permit consequence. The phrase contains the
        # words "permit is required", but its meaning is that the requirement has
        # NOT been established.
        epistemic_permit_finding = is_epistemic_permit_finding(permit_finding)
        if finding_claims_permit and not epistemic_permit_finding and not valid_permit_ids:
            errors.append(f"{discipline}: permit finding contains a permit consequence that is not established by valid PERMIT_REQUIREMENT evidence.")
        if permit == "VERIFIED_REQUIRED" and not valid_permit_ids:
            errors.append(f"{discipline}: VERIFIED_REQUIRED permit cannot survive without valid PERMIT_REQUIREMENT evidence.")
        if has_definitive_negative_permit_claim(permit_finding) and permit in {"NOT_APPLICABLE", "VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED"} and not valid_exemption_ids:
            errors.append(f"{discipline}: negative/exemption permit statement lacks valid PERMIT_EXEMPTION evidence.")
        if pathway == "VERIFIED_REQUIRED" and not valid_pathway_ids:
            errors.append(f"{discipline}: VERIFIED_REQUIRED pathway cannot survive without valid PATHWAY evidence.")

        # A model must not claim an established pathway while simultaneously
        # admitting that no proposition-specific rule was established. This
        # contradiction can otherwise slip through when a broad department
        # page contains generic process language (e.g. "plan check") but
        # does not actually establish the claimed requirement.
        applicability = item.get("applicability") or {}
        applicability_text = _norm_text(
            " ".join([
                str(applicability.get("rule") or ""),
                str(applicability.get("rule_finding") or ""),
                str(applicability.get("source_finding") or ""),
            ])
        ) if isinstance(applicability, dict) else ""
        no_specific_rule = bool(re.search(
            r"\bno\s+proposition[- ]specific\s+(?:authoritative\s+)?rule\s+was\s+established\b|"
            r"\bno\s+proposition[- ]specific\s+(?:authoritative\s+)?evidence\s+was\s+established\b",
            applicability_text, re.I
        ))
        if pathway == "VERIFIED_REQUIRED" and no_specific_rule:
            errors.append(
                f"{discipline}: pathway is marked VERIFIED_REQUIRED/Established even though the applicability section says no proposition-specific authoritative rule or evidence was established."
            )
        pathway_claim_patterns = [
            r"\bsubmit\b", r"\bapplication\b", r"\bportal\b", r"\bover[- ]the[- ]counter\b",
            r"\bplan review\b", r"\bpathway\b", r"\bprocessed\b", r"\bfile\b",
        ]
        if pathway and pathway != "UNKNOWN" and any(re.search(p, pathway_finding) for p in pathway_claim_patterns):
            if not valid_pathway_ids and pathway_basis == "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: pathway finding claims a specific process but has no valid PATHWAY evidence.")

    return errors

def validate_bottom_line(data):
    errors = []
    bottom_line = (data.get("bottom_line") or "").lower()
    if has_definitive_negative_permit_claim(bottom_line):
        errors.append("Bottom Line contains a definitive negative permit conclusion that requires explicit supporting evidence.")
    bottom_line_regulatory_markers = [r"\blb\b", r"\blbs\b", r"\bcfm\b", r"\bton(?:s)?\b", r"\bbtu\b", r"square feet", r"sq\.?\s*ft\.?", r"\bsection\b", r"\bchapter\b", r"\bthreshold\b", r"over-the-counter", r"minor label", r"full plan review", r"trade permit", r"administrative review", r"cup amendment", r"cup modification", r"energy permit"]
    if any(re.search(marker, bottom_line) for marker in bottom_line_regulatory_markers):
        if not data.get("bottom_line_evidence"):
            errors.append("Bottom Line contains a specific regulatory claim without supporting evidence IDs.")
    evidence_by_id = {e.get("id"): e for e in (data.get("evidence") or []) if e.get("id")}
    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        app = item.get("applicability", {})
        determination = app.get("determination")
        if permit == "VERIFIED_REQUIRED":
            matched_permit_ids = [
                eid for eid in (data.get("bottom_line_evidence") or [])
                if evidence_supports_permit_requirement(evidence_by_id.get(eid), discipline)
            ]
            if not matched_permit_ids:
                errors.append(f"Bottom Line: VERIFIED_REQUIRED {discipline} permit is not traced to discipline-matched PERMIT_REQUIREMENT evidence.")
        unresolved = (
            permit in {"CONDITIONAL", "UNKNOWN", "NOT_CURRENTLY_TRIGGERED"}
            or determination == "cannot_determine"
        )
        if not unresolved: continue
        discipline_words = [discipline.lower(), discipline.lower().replace("/", " ")]
        if not any(word in bottom_line for word in discipline_words if word): continue
        risky_patterns = [f"{discipline.lower()} permit is required", f"{discipline.lower()} permit required", 
                          f"{discipline.lower()} approval is required", f"{discipline.lower()} approval required",
                          f"{discipline.lower()} is not currently triggered", f"{discipline.lower()} not currently triggered"]
        for pattern in risky_patterns:
            if pattern in bottom_line:
                errors.append(f"Bottom Line may overstate {discipline}: discipline remains unresolved/conditional.")
                break
    return errors

# ============================================================
# GEMINI USAGE LOGGING
# ============================================================
USAGE_LOG_PATH = os.getenv("GEMINI_USAGE_LOG", "gemini_usage.log")

usage_logger = logging.getLogger("gemini_usage")
usage_logger.setLevel(logging.INFO)
usage_logger.propagate = False
if not usage_logger.handlers:
    try:
        _usage_handler = logging.FileHandler(USAGE_LOG_PATH, encoding="utf-8")
        _usage_handler.setFormatter(logging.Formatter("%(message)s"))
        usage_logger.addHandler(_usage_handler)
    except Exception:
        # Logging must never prevent the research run from executing.
        pass

def record_gemini_usage(model, response, caller="", prompt_hash="", elapsed_seconds=None):
    """Record one actual Gemini API response for per-request cost analysis."""
    meta = getattr(response, "usage_metadata", None)
    now = datetime.now(timezone.utc).isoformat()
    entry = {
        "timestamp_utc": now,
        "caller": caller,
        "model": model,
        "prompt_hash": prompt_hash,
        "elapsed_seconds": round(float(elapsed_seconds), 3) if elapsed_seconds is not None else None,
        "input_tokens": getattr(meta, "prompt_token_count", None) if meta else None,
        "output_tokens": getattr(meta, "candidates_token_count", None) if meta else None,
        "thoughts_tokens": getattr(meta, "thoughts_token_count", None) if meta else None,
        "total_tokens": getattr(meta, "total_token_count", None) if meta else None,
        "cached_tokens": getattr(meta, "cached_content_token_count", None) if meta else None,
    }
    try:
        usage_logger.info(json.dumps(entry, separators=(",", ":"), default=str))
    except Exception:
        pass
    print(f"[GEMINI USAGE] {json.dumps(entry, separators=(",", ":"), default=str)}")
    return entry

def record_gemini_exception(model, caller="", prompt_hash="", elapsed_seconds=None, error=""):
    """Record an API attempt that failed before a response with usage metadata existed."""
    entry = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "caller": caller,
        "model": model,
        "prompt_hash": prompt_hash,
        "elapsed_seconds": round(float(elapsed_seconds), 3) if elapsed_seconds is not None else None,
        "error": str(error)[:500],
    }
    try:
        usage_logger.info(json.dumps(entry, separators=(",", ":"), default=str))
    except Exception:
        pass
    print(f"[GEMINI ERROR] {json.dumps(entry, separators=(",", ":"), default=str)}")
    return entry

# ============================================================
# CACHING & API CALL
# ============================================================
# Streamlit's generic cache can cache a returned error object. That made a
# rejected Gemini request (including a 403 spend-cap error) look like a
# permanent no-op on subsequent clicks. Cache only successful API results.
def cache_success_only(ttl=3600):
    def decorator(func):
        def wrapper(*args, **kwargs):
            cache = st.session_state.setdefault("gemini_success_cache", {})
            key_payload = {"function": func.__name__, "args": args, "kwargs": kwargs}
            key = hashlib.md5(json.dumps(key_payload, sort_keys=True, default=str).encode()).hexdigest()
            now = time.time()
            cached = cache.get(key)
            if isinstance(cached, dict) and now - float(cached.get("stored_at", 0)) < ttl:
                return copy.deepcopy(cached.get("result"))
            result = func(*args, **kwargs)
            if isinstance(result, dict) and not result.get("error"):
                cache[key] = {"stored_at": now, "result": copy.deepcopy(result)}
            return result
        wrapper.__name__ = getattr(func, "__name__", "cached_function")
        wrapper.__doc__ = getattr(func, "__doc__", None)
        return wrapper
    return decorator
@cache_success_only(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": 1}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=prompt_text,
            config=config,
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="initial_research",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        if response.candidates:
            candidate = response.candidates[0]
            finish_reason = str(candidate.finish_reason)
            debug_info["finish_reason"] = finish_reason
            ratings = getattr(candidate, "safety_ratings", None) or []
            debug_info["safety_ratings"] = [str(r) for r in ratings]
            if finish_reason == "FinishReason.MAX_TOKENS":
                debug_info["error_type"] = "MAX_TOKENS"
                return {"data": None, "error": True, "retry": True, "msg": "First research pass reached the output limit.", "debug": debug_info}
            if finish_reason and finish_reason != "FinishReason.STOP":
                debug_info["error_type"] = "Early Stop"
                return {"data": None, "error": True, "retry": False, "msg": f"API stopped early: {finish_reason}", "debug": debug_info}
        else:
            debug_info["candidates"] = None
            debug_info["prompt_feedback"] = str(getattr(response, "prompt_feedback", None))
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "retry": False, "msg": "No candidates returned.", "debug": debug_info}
        text = getattr(response, "text", None)
        if not text:
            try: text = response.candidates[0].content.parts[0].text
            except Exception: pass
        if not text:
            debug_info["raw_text"] = ""
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "retry": False, "msg": "Empty response.", "debug": debug_info}
        debug_info["text_length"] = len(text)
        try:
            data = extract_json(text)
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "error": True, "retry": False, "msg": "Failed to parse JSON.", "debug": debug_info}
        
        data = normalize_dossier_status_values(data)
        data = sanitize_scope_relevance(data)
        data = sanitize_expected_process(data)
        data = normalize_dossier_basis_values(data)
        data = sanitize_unsupported_not_currently_triggered_statuses(data)
        data = sanitize_unverifiable_verified_jurisdiction(data, str(address))
        try:
            _as_of = datetime.strptime(str(project_date), "%Y-%m-%d").date()
        except Exception:
            _as_of = date.today()
        data = sanitize_invalid_current_codes(data, _as_of)
        data = sanitize_invalid_evidence_propositions(data)
        data = sanitize_contradictory_established_pathways(data)
        data = sanitize_generic_proposition_links(data)
        data = sanitize_invalid_authority_evidence_links(data)
        data = sanitize_cross_discipline_applicability_links(data) # ADDED HERE
        data = repair_missing_threshold_links(data)
        data = sanitize_unsupported_permit_conclusions(data)
        data = sanitize_unsupported_threshold_conclusions(data)
        data = sanitize_unsupported_entitlement_rules(data)
        data = sanitize_unsupported_entitlement_conclusions(data)
        data = sanitize_unsupported_pathway_conclusions(data)
        data = sanitize_unverifiable_verified_pathways(data)
        data = sanitize_unverifiable_verified_permits(data)
        data = sanitize_semantically_misplaced_permit_findings(data)
        data = sanitize_unsubstantiated_conditional_statuses(data)
        data = sanitize_bottom_line_for_jurisdiction(data)
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_final_regulatory_statuses(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_validation_blocking_permit_findings(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)
        
        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            debug_info["error_type"] = "Validation Failed"
            return {
                "data": data,
                "error": True,
                "retry": False,
                "msg": "Research completed, but the dossier failed regulatory consistency validation.",
                "debug": debug_info,
            }
        debug_info["status"] = "success"
        return {"data": data, "error": False, "retry": False, "debug": debug_info}
    except Exception as e:
        error_msg = str(e)
        debug_info["exception"] = error_msg
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="initial_research",
            prompt_hash=prompt_hash, elapsed_seconds=(time.monotonic() - _api_started) if "_api_started" in locals() else None,
            error=error_msg
        )
        debug_info["error_type"] = "Python Exception"
        if "spend cap breached" in error_msg.lower():
            debug_info["error_type"] = "SPEND_CAP_BREACHED"
            return {"data": None, "error": True, "retry": False, "msg": "Google Gemini rejected this request because the project spend cap has been breached. No Gemini generation completed and no normal token usage should be recorded for this attempt.", "debug": debug_info}
        if "429" in error_msg:
            debug_info["error_type"] = "QUOTA_EXCEEDED"
            return {"data": None, "error": True, "retry": False, "msg": "Google Gemini rejected the request because the project quota/rate limit was exceeded.", "debug": debug_info}
        if "403" in error_msg and "permission" in error_msg.lower():
            debug_info["error_type"] = "PERMISSION_DENIED"
            return {"data": None, "error": True, "retry": False, "msg": "Google Gemini rejected the request with 403 PERMISSION_DENIED. See API Debug Log for the exact Google message.", "debug": debug_info}
        return {"data": None, "error": True, "retry": False, "msg": f"Gemini API error: {error_msg[:300]}", "debug": debug_info}


@cache_success_only(ttl=3600)
def cached_gemini_jurisdiction_recovery(prompt_hash, recovery_prompt):
    """Run a focused paid pass to establish the parcel's actual governmental jurisdiction."""
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": "jurisdiction_recovery"}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=8192,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash", contents=recovery_prompt, config=config
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="jurisdiction_recovery",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        if not response.candidates:
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "msg": "Jurisdiction recovery returned no candidates.", "debug": debug_info}
        finish_reason = str(response.candidates[0].finish_reason)
        debug_info["finish_reason"] = finish_reason
        if finish_reason == "FinishReason.MAX_TOKENS":
            debug_info["error_type"] = "MAX_TOKENS"
            return {"data": None, "error": True, "msg": "Jurisdiction recovery reached the output limit.", "debug": debug_info}
        if finish_reason and finish_reason != "FinishReason.STOP":
            debug_info["error_type"] = "Early Stop"
            return {"data": None, "error": True, "msg": f"Jurisdiction recovery stopped early: {finish_reason}", "debug": debug_info}
        text = getattr(response, "text", None) or ""
        if not text:
            try:
                text = response.candidates[0].content.parts[0].text
            except Exception:
                text = ""
        if not text:
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "msg": "Jurisdiction recovery returned empty text.", "debug": debug_info}
        debug_info["text_length"] = len(text)
        try:
            data = extract_json(text)
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "msg": "Jurisdiction recovery JSON could not be parsed.", "debug": debug_info}
        debug_info["status"] = "success"
        return {"data": data, "error": False, "debug": debug_info}
    except Exception as e:
        error_msg = str(e)
        debug_info["error_type"] = "Python Exception"
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="jurisdiction_recovery", prompt_hash=prompt_hash,
            elapsed_seconds=(time.monotonic() - _api_started) if "_api_started" in locals() else None,
            error=error_msg
        )
        return {"data": None, "error": True, "msg": f"Jurisdiction recovery error: {error_msg[:200]}", "debug": debug_info}


def build_jurisdiction_recovery_prompt(data, address, state, project_date):
    return f"""
You are conducting a TARGETED PARCEL-JURISDICTION VERIFICATION PASS. Return ONLY valid JSON.

PROJECT: State={state} | Address={address} | Date={project_date}

The initial research did not establish the parcel's actual governmental boundary with sufficient
confidence. Determine the governmental jurisdiction that legally governs THIS PARCEL.

HARD RULES:
- Do NOT treat the postal city, mailing city, ZIP code, or address city as proof that the parcel is
  inside that municipality.
- Do NOT infer municipal jurisdiction from a city agency page merely because the address appears on it.
- Prefer an official county parcel/GIS/assessor/property record, official municipal boundary/GIS source,
  official planning jurisdiction map, or other governmental record that expressly establishes whether the
  parcel is inside city limits or unincorporated county jurisdiction.
- The evidence rule must state the actual parcel/boundary result, not merely that an agency serves the area.
- If the parcel cannot be verified, return no jurisdiction evidence rather than guessing.
- Do not research permits or process in this pass.
- Do not invent URLs, parcel numbers, boundary results, section numbers, or agency relationships.

OUTPUT:
{{
  "jurisdiction": {{"status":"VERIFIED|CONDITIONAL", "county":"string", "city":"string", "ahj":"string", "evidence":["J1"]}},
  "evidence": [{{
    "id":"J1", "title":"actual source title", "url":"actual source URL",
    "authority":"county|city|state|other", "discipline":"Jurisdiction",
    "proposition_type":"JURISDICTION", "source_type":"other|code|ordinance",
    "retrieval_note":"short", "rule":"explicit parcel/boundary jurisdiction proposition"
  }}]
}}
"""

@cache_success_only(ttl=3600)
def cached_gemini_permit_recovery(prompt_hash, recovery_prompt):
    """Run a focused second research pass for unresolved permit consequences."""
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": "permit_recovery"}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=16384,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash", contents=recovery_prompt, config=config
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="permit_recovery",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        if not response.candidates:
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "msg": "Permit recovery returned no candidates.", "debug": debug_info}
        finish_reason = str(response.candidates[0].finish_reason)
        debug_info["finish_reason"] = finish_reason
        if finish_reason == "FinishReason.MAX_TOKENS":
            debug_info["error_type"] = "MAX_TOKENS"
            return {"data": None, "error": True, "msg": "Permit recovery reached the output limit.", "debug": debug_info}
        if finish_reason and finish_reason != "FinishReason.STOP":
            debug_info["error_type"] = "Early Stop"
            return {"data": None, "error": True, "msg": f"Permit recovery stopped early: {finish_reason}", "debug": debug_info}
        text = getattr(response, "text", None) or ""
        if not text:
            try:
                text = response.candidates[0].content.parts[0].text
            except Exception:
                text = ""
        if not text:
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "msg": "Permit recovery returned empty text.", "debug": debug_info}
        debug_info["text_length"] = len(text)
        try:
            data = extract_json(text)
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "msg": "Permit recovery JSON could not be parsed.", "debug": debug_info}
        debug_info["status"] = "success"
        return {"data": data, "error": False, "debug": debug_info}
    except Exception as e:
        error_msg = str(e)
        debug_info["error_type"] = "Python Exception"
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="permit_recovery", prompt_hash=prompt_hash,
            elapsed_seconds=(time.monotonic() - _api_started) if "_api_started" in locals() else None,
            error=error_msg
        )
        return {"data": None, "error": True, "msg": f"Permit recovery error: {error_msg[:200]}", "debug": debug_info}


def merge_jurisdiction_recovery(data, recovery, address=""):
    """Merge only validated, site-specific jurisdiction evidence from recovery."""
    if not isinstance(data, dict) or not isinstance(recovery, dict):
        return data
    incoming = recovery.get("evidence") or []
    if isinstance(incoming, dict):
        incoming = [incoming]
    existing = data.get("evidence") or []
    if not isinstance(existing, list):
        existing = []
    existing_ids = {str(e.get("id")) for e in existing if isinstance(e, dict) and e.get("id")}
    new_ids = []
    for ev in incoming:
        if not isinstance(ev, dict):
            continue
        if ev.get("proposition_type") != "JURISDICTION":
            continue
        if jurisdiction_evidence_integrity_errors(ev) or evidence_proposition_integrity_errors(ev):
            continue
        if not _jurisdiction_evidence_is_site_specific(ev, address):
            continue
        eid = str(ev.get("id") or "").strip()
        if not eid:
            continue
        if eid in existing_ids:
            continue
        existing.append(ev)
        existing_ids.add(eid)
        new_ids.append(eid)
    if new_ids:
        data["evidence"] = existing
        j = data.setdefault("jurisdiction", {})
        rj = recovery.get("jurisdiction") or {}
        j["status"] = "VERIFIED"
        if rj.get("county"):
            j["county"] = rj.get("county")
        if rj.get("city"):
            j["city"] = rj.get("city")
        if rj.get("ahj"):
            j["ahj"] = rj.get("ahj")
        j["evidence"] = list(dict.fromkeys([*(j.get("evidence") or []), *new_ids]))
    return data


def sanitize_unverified_jurisdiction_identity(data):
    """Do not present a postal city/AHJ as the governing jurisdiction when boundary verification is absent."""
    if not isinstance(data, dict):
        return data
    j = data.get("jurisdiction") or {}
    if str(j.get("status") or "").upper() == "VERIFIED":
        return data
    # County can remain as a research candidate; municipal identity and AHJ cannot.
    j["city"] = "Not yet established"
    j["ahj"] = "Unconfirmed"
    return data

def unresolved_permit_recovery_targets(data):
    """Return disciplines whose permit consequence still lacks a validated answer.

    The initial Gemini response is allowed to use aliases such as NOT_ESTABLISHED,
    INSUFFICIENT_EVIDENCE, or UNKNOWN_EVIDENCE.  Recovery must normalize those aliases
    before deciding whether a second research pass is needed; otherwise a dossier can
    reach the final report unresolved without ever making the paid recovery call.
    """
    targets = []
    if not isinstance(data, dict):
        return targets

    normalized = normalize_dossier_status_values(data)
    for item in normalized.get("disciplines") or []:
        if not isinstance(item, dict):
            continue

        permit = re.sub(r"[^A-Z0-9_]+", "_", str(item.get("permit") or "").strip().upper()).strip("_")
        basis = str(item.get("permit_basis") or "").strip().upper()
        evidence = item.get("permit_evidence") or []

        # Anything without a validated permit decision is a recovery target.
        # A definitive status is not enough by itself; the deterministic evidence
        # contract remains the source of truth.
        unresolved = (
            permit in {"CONDITIONAL", "UNKNOWN", "NOT_ESTABLISHED", "INSUFFICIENT_EVIDENCE", "INSUFFICIENT", "UNDETERMINED"}
            or basis == "NOT_ESTABLISHED"
            or not evidence
        )
        if not unresolved:
            continue

        app = item.get("applicability") or {}
        fact = app.get("fact") or {}
        targets.append({
            "type": str(item.get("type") or "Unknown"),
            "applicability": str(app.get("rule") or ""),
            "fact": str(fact.get("statement") or ""),
            "permit_finding": str(item.get("permit_finding") or ""),
        })
    return targets


def merge_permit_recovery(data, recovery):
    """Merge only validated permit evidence from the focused recovery pass.

    Gemini does not get to decide permit status, pathway, authority, or bottom-line
    conclusions here.  Recovery contributes source records only; the deterministic
    contract derives the permit status from those records afterward.
    """
    if not isinstance(data, dict) or not isinstance(recovery, dict):
        return data

    evidence_by_id = {
        str(e.get("id")): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    accepted_ids = set()
    recovery_id_map = {}
    next_index = 1

    for ev in recovery.get("evidence") or []:
        if not isinstance(ev, dict):
            continue
        ptype = str(ev.get("proposition_type") or "").strip().upper()
        discipline = str(ev.get("discipline") or "").strip()

        # This pass is strictly for permit consequences.  Do not allow Gemini to
        # smuggle pathway, entitlement, jurisdiction, or code conclusions through it.
        if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION"}:
            continue
        if not discipline:
            continue

        candidate = dict(ev)
        candidate["proposition_type"] = ptype
        old_id = str(candidate.get("id") or "").strip()
        if not old_id:
            continue

        # Validate the evidence before it enters the main dossier.  The same
        # proposition/source contract used by the final validator is enforced here.
        if evidence_proposition_integrity_errors(candidate):
            continue
        if evidence_source_specificity_errors([candidate]):
            continue

        # Namespace recovery IDs so they cannot collide with the primary pass.
        while True:
            new_id = f"PR{next_index}"
            next_index += 1
            if new_id not in evidence_by_id:
                break
        candidate["id"] = new_id
        recovery_id_map[old_id] = new_id
        evidence_by_id[new_id] = candidate
        accepted_ids.add(new_id)

    if not accepted_ids:
        return data

    data["evidence"] = list(evidence_by_id.values())

    # Attach only the validated recovery evidence to the matching discipline.
    # Do not accept Gemini-supplied status/finding/basis/pathway fields.
    disciplines = {
        str(item.get("type") or "").strip().lower(): item
        for item in data.get("disciplines") or []
        if isinstance(item, dict)
    }
    for upd in recovery.get("permit_updates") or recovery.get("discipline_updates") or []:
        if not isinstance(upd, dict):
            continue
        item = disciplines.get(str(upd.get("type") or "").strip().lower())
        if not item:
            continue
        raw_ids = upd.get("permit_evidence") or []
        if isinstance(raw_ids, str):
            raw_ids = [raw_ids]
        if not isinstance(raw_ids, list):
            continue
        mapped = [recovery_id_map.get(str(eid)) for eid in raw_ids]
        mapped = [eid for eid in mapped if eid in accepted_ids]
        if not mapped:
            continue

        existing = item.get("permit_evidence") or []
        if isinstance(existing, str):
            existing = [existing]
        if not isinstance(existing, list):
            existing = []
        item["permit_evidence"] = list(dict.fromkeys(existing + mapped))

    # Some recovery implementations may return evidence without a separate update.
    # Attach it by discipline so useful evidence is not discarded.
    for new_id in accepted_ids:
        ev = evidence_by_id.get(new_id) or {}
        discipline = str(ev.get("discipline") or "").strip().lower()
        item = disciplines.get(discipline)
        if not item:
            continue
        existing = item.get("permit_evidence") or []
        if isinstance(existing, str):
            existing = [existing]
        if not isinstance(existing, list):
            existing = []
        if new_id not in existing:
            item["permit_evidence"] = existing + [new_id]

    return data

def unresolved_process_recovery_targets(data):
    """Return disciplines where the plan-review/submittal process is still unresolved."""
    targets = []
    if not isinstance(data, dict):
        return targets
    evidence_by_id = {
        str(e.get("id")): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines") or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown").strip()
        pathway = str(item.get("pathway") or "UNKNOWN").strip().upper()
        pathway_basis = str(item.get("pathway_basis") or "").strip().upper()
        ids = item.get("pathway_evidence") or []
        if isinstance(ids, str):
            ids = [ids]
        valid_pathway = any(
            evidence_supports_pathway(evidence_by_id.get(str(eid)), discipline)
            or evidence_supports_review(evidence_by_id.get(str(eid)), discipline)
            for eid in ids
        )
        unresolved = (
            pathway in {"CONDITIONAL", "UNKNOWN", "NOT_ESTABLISHED", "INSUFFICIENT", "INSUFFICIENT_EVIDENCE", "UNDETERMINED"}
            or pathway_basis != "DIRECT_EVIDENCE"
            or not valid_pathway
        )
        if not unresolved:
            continue
        app = item.get("applicability") or {}
        fact = app.get("fact") or {}
        targets.append({
            "type": discipline,
            "permit": str(item.get("permit") or "UNKNOWN"),
            "scope_fact": str(fact.get("statement") or ""),
            "pathway_finding": str(item.get("pathway_finding") or ""),
        })
    return targets


@cache_success_only(ttl=3600)
def cached_gemini_process_recovery(prompt_hash, recovery_prompt):
    """Focused paid pass for who reviews the work and how each discipline is processed."""
    debug_info = {"status": "processing", "attempt": 1, "caller": "process_recovery"}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=16384,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash", contents=recovery_prompt, config=config
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="process_recovery",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        if not response.candidates:
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "msg": "Process recovery returned no candidates.", "debug": debug_info}
        finish_reason = str(response.candidates[0].finish_reason)
        debug_info["finish_reason"] = finish_reason
        if finish_reason == "FinishReason.MAX_TOKENS":
            debug_info["error_type"] = "MAX_TOKENS"
            return {"data": None, "error": True, "msg": "Process recovery reached the output limit.", "debug": debug_info}
        if finish_reason and finish_reason != "FinishReason.STOP":
            debug_info["error_type"] = "Early Stop"
            return {"data": None, "error": True, "msg": f"Process recovery stopped early: {finish_reason}", "debug": debug_info}
        text = getattr(response, "text", None) or ""
        if not text:
            try:
                text = response.candidates[0].content.parts[0].text
            except Exception:
                text = ""
        if not text:
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "msg": "Process recovery returned empty text.", "debug": debug_info}
        debug_info["text_length"] = len(text)
        try:
            data = extract_json(text)
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "msg": "Process recovery JSON could not be parsed.", "debug": debug_info}
        debug_info["status"] = "success"
        return {"data": data, "error": False, "debug": debug_info}
    except Exception as e:
        error_msg = str(e)
        debug_info["error_type"] = "Python Exception"
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="process_recovery", prompt_hash=prompt_hash,
            elapsed_seconds=(time.monotonic() - _api_started) if "_api_started" in locals() else None,
            error=error_msg
        )
        return {"data": None, "error": True, "msg": f"Process recovery error: {error_msg[:200]}", "debug": debug_info}


def merge_process_recovery(data, recovery):
    """Merge only validated process evidence; Gemini cannot directly set pathway status."""
    if not isinstance(data, dict) or not isinstance(recovery, dict):
        return data
    evidence_by_id = {
        str(e.get("id")): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    recovery_id_map = {}
    accepted = {}
    next_index = 1
    for ev in recovery.get("evidence") or []:
        if not isinstance(ev, dict):
            continue
        ptype = str(ev.get("proposition_type") or "").strip().upper()
        if ptype not in {"PATHWAY", "REVIEW_REQUIREMENT", "AUTHORITY_HIERARCHY"}:
            continue
        candidate = dict(ev)
        old_id = str(candidate.get("id") or "").strip()
        if not old_id:
            continue
        if evidence_proposition_integrity_errors(candidate):
            continue
        if ptype in {"PATHWAY", "REVIEW_REQUIREMENT"} and not str(candidate.get("url") or "").strip():
            continue
        if ptype == "PATHWAY" and evidence_source_specificity_errors([candidate]):
            continue
        if ptype == "REVIEW_REQUIREMENT" and _is_generic_review_source(candidate):
            continue
        if _jurisdiction_process_source_mismatch(candidate, data.get("jurisdiction") or {}):
            continue
        while True:
            new_id = f"PRC{next_index}"
            next_index += 1
            if new_id not in evidence_by_id:
                break
        candidate["id"] = new_id
        evidence_by_id[new_id] = candidate
        accepted[new_id] = candidate
        recovery_id_map[old_id] = new_id
    if not accepted:
        return data
    data["evidence"] = list(evidence_by_id.values())
    disciplines = {str(i.get("type") or "").strip().lower(): i for i in data.get("disciplines") or [] if isinstance(i, dict)}
    for upd in recovery.get("process_updates") or []:
        if not isinstance(upd, dict):
            continue
        item = disciplines.get(str(upd.get("type") or "").strip().lower())
        if not item:
            continue
        raw = upd.get("pathway_evidence") or upd.get("review_evidence") or upd.get("process_evidence") or []
        if isinstance(raw, str):
            raw = [raw]
        mapped = [recovery_id_map.get(str(x)) for x in raw]
        mapped = [x for x in mapped if x in accepted]
        if mapped:
            existing = item.get("process_evidence") or []
            if isinstance(existing, str):
                existing = [existing]
            item["process_evidence"] = list(dict.fromkeys(existing + mapped))
            # Keep legacy pathway evidence separate. Process recovery must not silently
            # turn a documented reviewer/process step into a claimed routing pathway.
            update_fields = {
                "reviewer": "process_reviewer",
                "route": "process_route",
                "relationship": "process_relationship",
                "trigger": "process_trigger",
            }
            for src_key, dst_key in update_fields.items():
                value = str(upd.get(src_key) or "").strip()
                if value:
                    item[dst_key] = value
    for new_id, ev in accepted.items():
        discipline = str(ev.get("discipline") or "").strip().lower()
        if discipline in disciplines and ev.get("proposition_type") in {"PATHWAY", "REVIEW_REQUIREMENT"}:
            existing = disciplines[discipline].get("process_evidence") or []
            if isinstance(existing, str):
                existing = [existing]
            if new_id not in existing:
                disciplines[discipline]["process_evidence"] = existing + [new_id]
        elif discipline in disciplines and ev.get("proposition_type") == "AUTHORITY_HIERARCHY":
            existing = disciplines[discipline].get("authority_evidence") or []
            if isinstance(existing, str):
                existing = [existing]
            if new_id not in existing:
                disciplines[discipline]["authority_evidence"] = existing + [new_id]
    return derive_process_status_from_evidence(data)


def build_process_recovery_prompt(data, address, state, project_date, ptype, bclass, sow_text, targets=None):
    if targets is None:
        targets = unresolved_process_recovery_targets(data)
    return f"""
You are conducting a SECOND TARGETED AHJ PROCESS RESEARCH PASS. Return ONLY valid JSON.
Recover the actual plan-review/submittal PROCESS, not another code applicability summary.

PROJECT: State={state} | Address={address} | Date={project_date}
Type={ptype} | Class={bclass}
SCOPE: {sow_text}

VERIFIED JURISDICTION CONTEXT:
{json.dumps(data.get("jurisdiction") or {}, indent=2)}
HARD JURISDICTION BOUNDARY:
- Research the governmental process that actually applies to this parcel, not the postal city process.
- If City is "Unincorporated", do NOT use a municipal city plan-check document merely because the
  postal address contains that city name. Use the verified County/AHJ process instead.
- If jurisdiction Status is not VERIFIED, do NOT use municipal process evidence as the governing route.
  Keep the process unresolved unless the source itself establishes the applicable governmental authority.
- A city source may be used only if the source itself explicitly establishes that the County uses or
  adopts that city process for unincorporated parcels. Otherwise reject it.
- The reviewer, route, relationship, and trigger must all be traceable to the verified jurisdiction.

TARGETS:
{json.dumps(targets, indent=2)}

For each target, determine as specifically as authoritative sources allow:
1. WHO performs the review/plan check (Building & Safety, Fire Department, Public Works,
   Planning, state agency, utility, third-party reviewer, etc.).
2. WHETHER that discipline has its own separate review, referral, approval, inspection,
   or plan-check step versus being reviewed as part of the primary building permit.
3. HOW the submittal is routed (permit application, plan check, e-permit portal, counter,
   referral, concurrent review, sequential review, or other documented route).
4. WHAT project fact, permit type, threshold, or document changes the process.
5. WHEN another discipline/AHJ is involved, if the source expressly establishes that relationship.

SOURCE PRIORITY:
1. Official AHJ permit/checklist/application/plan-review instructions.
2. Official local ordinance/code section expressly describing review/referral/process.
3. Official AHJ FAQ or workflow document.
4. Official interagency agreement or governmental guidance.

HARD CONTRACT:
- Evidence must state the actual process/reviewer proposition. Code applicability alone is not process evidence.
- A generic homepage, department landing page, code index, or portal homepage is not process evidence.
- Do not infer who reviews something merely because an agency has jurisdiction over that subject.
- Do not infer that every discipline gets a separate plan check. The source must establish the separate/concurrent/referral relationship.
- Do not infer a sequence unless the source states or clearly documents it.
- If the process cannot be established, return no evidence for that target.
- Do not return pathway status, findings, basis, permit status, bottom line, or other conclusions.
- Do not invent URLs, section numbers, titles, or process steps.
- The evidence URL must be the actual source URL from the grounded web result, never a vertexaisearch.cloud.google.com grounding redirect.
- Only populate reviewer, route, relationship, or trigger fields when the cited evidence explicitly supports that field.
- If a field is not established by the source, return an empty string for that field.

OUTPUT:
{{
  "evidence": [{{
    "id":"R1", "title":"actual source title", "url":"actual source URL",
    "authority":"state|county|city|federal|tribal|other",
    "discipline":"exact existing discipline type",
    "proposition_type":"PATHWAY|REVIEW_REQUIREMENT|AUTHORITY_HIERARCHY",
    "source_type":"permit_page|checklist|application|ordinance|code|interpretation|other",
    "retrieval_note":"short",
    "rule":"explicit process/reviewer proposition stated by source"
  }}],
  "process_updates": [{{
    "type":"exact existing discipline type",
    "process_evidence":["R1"],
    "reviewer":"who actually performs the documented review, only if source states it",
    "route":"how the work is submitted/routed, only if source states it",
    "relationship":"separate/concurrent/referral/part-of-primary-review, only if source states it",
    "trigger":"project fact or document that changes this process, only if source states it"
  }}]
}}
"""

def build_permit_recovery_prompt(data, address, state, project_date, ptype, bclass, sow_text, targets=None):
    if targets is None:
        targets = unresolved_permit_recovery_targets(data)
    return f"""
You are conducting a SECOND TARGETED AHJ RESEARCH PASS. Return ONLY valid JSON.
Do not rewrite the dossier. Recover missing proposition-specific permit evidence.

PROJECT: State={state} | Address={address} | Date={project_date}
Type={ptype} | Class={bclass}
SCOPE: {sow_text}

UNRESOLVED TARGETS:
{json.dumps(targets, indent=2)}

For EACH target, search the verified AHJ and controlling governmental sources specifically
for the LEGAL PERMIT OR APPROVAL CONSEQUENCE of the actual work. This is not another
code-applicability search.

SEARCH TERMS: combine the actual scope item with permit language. Examples:
commercial panel replacement permit; commercial plumbing fixture replacement permit;
HVAC replacement permit; roof replacement permit; structural repair permit; window infill
permit; fire alarm permit; accessibility alteration permit; energy compliance permit;
zoning clearance; encroachment permit; sewer connection permit. Adapt to the actual AHJ.

SOURCE PRIORITY:
1. Official AHJ permit requirement/checklist/application.
2. Official local ordinance or adopted code section expressly requiring the permit/approval.
3. Official AHJ fee schedule/application instruction expressly identifying the permit.
4. Official AHJ FAQ/interpretation expressly stating the requirement.

HARD SOURCE CONTRACT:
- Generic homepages, department landing pages, permit portals, code indexes, and code-adoption
  pages are NOT permit evidence unless the page itself contains the claimed proposition.
- The evidence rule itself must explicitly state that the permit/approval is required, needed,
  necessary, or must be obtained.
- Code applicability, compliance requirements, inspection, plan review, application existence,
  portal existence, regulated work, thresholds, and common construction practice do NOT by
  themselves establish a permit requirement.
- Never infer a permit merely because commercial work is regulated.
- If no explicit permit rule can be found, return no permit evidence for that target. That is
  an acceptable result. Keep the target unresolved.
- Do NOT write a permit-required conclusion together with language such as "no proposition-specific
  authoritative rule/evidence was established." If the rule was not established, the permit must
  remain unresolved.
- An explicit exemption is PERMIT_EXEMPTION only when the source itself says the permit is
  not required or an exemption applies.
- State-level permit rules require authority-hierarchy support when local control could modify
  the result.
- Do not invent URLs, section titles, quotations, or permit names.

OUTPUT — EVIDENCE ONLY:
{{
  "evidence": [{{
    "id":"R1", "title":"actual source title", "url":"actual source URL",
    "authority":"state|county|city|federal|tribal|other",
    "discipline":"exact existing discipline type",
    "proposition_type":"PERMIT_REQUIREMENT|PERMIT_EXEMPTION",
    "source_type":"code|ordinance|permit_page|checklist|application|interpretation|other",
    "retrieval_note":"short", "rule":"explicit proposition stated by source"
  }}],
  "permit_updates": [{{
    "type":"exact existing discipline type",
    "permit_evidence":["R1"]
  }}]
}}

IMPORTANT: Do NOT return permit status, permit_finding, permit_basis, pathway,
pathway_finding, pathway_basis, authority_evidence, bottom_line, or any other conclusion.
The application will derive the permit status deterministically from validated evidence.
Only include an evidence record when the source itself expressly establishes the permit
requirement or exemption. If no such source is found, omit the target.
"""

@cache_success_only(ttl=3600)
def cached_gemini_retry(prompt_hash, retry_prompt):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": 2}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=retry_prompt,
            config=config,
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="generation_retry",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        if not response.candidates:
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "retry": False, "msg": "Retry returned no candidates.", "debug": debug_info}
        candidate = response.candidates[0]
        finish_reason = str(candidate.finish_reason)
        debug_info["finish_reason"] = finish_reason
        if finish_reason == "FinishReason.MAX_TOKENS":
            debug_info["error_type"] = "MAX_TOKENS_RETRY"
            return {
                "data": None,
                "error": True,
                "retry": False,
                "msg": "Gemini reached the generation limit twice. The research request was too large for the current generation budget.",
                "debug": debug_info,
            }
        if finish_reason and finish_reason != "FinishReason.STOP":
            debug_info["error_type"] = "Early Stop"
            return {"data": None, "error": True, "retry": False, "msg": f"Retry stopped early: {finish_reason}", "debug": debug_info}
        text = getattr(response, "text", None)
        if not text:
            try: text = response.candidates[0].content.parts[0].text
            except Exception: text = None
        if not text:
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "retry": False, "msg": "Retry returned empty text.", "debug": debug_info}
        debug_info["text_length"] = len(text)
        try:
            data = extract_json(text)
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "retry": False, "msg": "Retry produced invalid JSON.", "debug": debug_info}
        
        data = normalize_dossier_status_values(data)
        data = sanitize_scope_relevance(data)
        data = sanitize_expected_process(data)
        data = normalize_dossier_basis_values(data)
        data = sanitize_unsupported_not_currently_triggered_statuses(data)
        data = sanitize_unverifiable_verified_jurisdiction(data, str(address))
        try:
            _as_of = datetime.strptime(str(project_date), "%Y-%m-%d").date()
        except Exception:
            _as_of = date.today()
        data = sanitize_invalid_current_codes(data, _as_of)
        data = sanitize_invalid_evidence_propositions(data)
        data = sanitize_generic_proposition_links(data)
        data = sanitize_invalid_authority_evidence_links(data)
        data = sanitize_cross_discipline_applicability_links(data) # ADDED HERE
        data = repair_missing_threshold_links(data)
        data = sanitize_unsupported_permit_conclusions(data)
        data = sanitize_unsupported_threshold_conclusions(data)
        data = sanitize_unsupported_entitlement_rules(data)
        data = sanitize_unsupported_entitlement_conclusions(data)
        data = sanitize_unsupported_pathway_conclusions(data)
        data = sanitize_unverifiable_verified_pathways(data)
        data = sanitize_unverifiable_verified_permits(data)
        data = sanitize_semantically_misplaced_permit_findings(data)
        data = sanitize_unsubstantiated_conditional_statuses(data)
        data = sanitize_bottom_line_for_jurisdiction(data)
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_final_regulatory_statuses(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_validation_blocking_permit_findings(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)
        
        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            debug_info["error_type"] = "Validation Failed"
            return {
                "data": data,
                "error": True,
                "retry": False,
                "msg": "Research completed, but the dossier failed regulatory consistency validation.",
                "debug": debug_info,
            }
        debug_info["status"] = "success"
        return {"data": data, "error": False, "retry": False, "debug": debug_info}
    except Exception as e:
        debug_info["exception"] = str(e)
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="generation_retry",
            prompt_hash=prompt_hash, elapsed_seconds=(time.monotonic() - _api_started) if "_api_started" in locals() else None,
            error=str(e)
        )
        debug_info["error_type"] = "Python Exception"
        return {"data": None, "error": True, "retry": False, "msg": f"Retry error: {str(e)[:200]}", "debug": debug_info}

# ============================================================
# VALIDATION-AWARE SELF-CORRECTION
# ============================================================


def constrain_validation_repair_output(repaired_data, previous_data):
    """Keep validation repair inside the original research/evidence boundary."""
    if not isinstance(repaired_data, dict):
        return previous_data if isinstance(previous_data, dict) else repaired_data
    if not isinstance(previous_data, dict):
        return repaired_data

    # Research facts are not repairable prose. Preserve them exactly.
    for key in (
        "evidence", "jurisdiction", "codes", "project", "project_facts",
        "scope", "address", "state", "project_date"
    ):
        if key in previous_data:
            repaired_data[key] = previous_data[key]

    # Research facts are not repairable structure either.  A validation repair
    # may change wording, but it must not rename/reorder/drop disciplines.
    # Otherwise a repair can turn "Building / Structural" into "Structural"
    # and break discipline-to-evidence matching even though the evidence itself
    # is unchanged.  Restore the canonical discipline identities/order while
    # retaining the repair's field-level corrections for each corresponding item.
    previous_disciplines = previous_data.get("disciplines") or []
    repaired_disciplines = repaired_data.get("disciplines") or []
    if isinstance(previous_disciplines, list) and isinstance(repaired_disciplines, list):
        repaired_by_family = {}
        for item in repaired_disciplines:
            if isinstance(item, dict):
                repaired_by_family.setdefault(discipline_family(item.get("type")), []).append(item)
        canonical = []
        for prev in previous_disciplines:
            if not isinstance(prev, dict):
                continue
            family = discipline_family(prev.get("type"))
            candidates = repaired_by_family.get(family) or []
            repaired_item = candidates.pop(0) if candidates else None
            if repaired_item is None:
                # Preserve the pre-repair item if Gemini omitted it.
                repaired_item = dict(prev)
            repaired_item["type"] = prev.get("type")
            canonical.append(repaired_item)
        repaired_data["disciplines"] = canonical

    # Never allow a repair response to introduce evidence IDs.
    previous_ids = {
        str(e.get("id")) for e in (previous_data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    repaired_evidence = repaired_data.get("evidence") or []
    if any(
        isinstance(e, dict) and e.get("id") and str(e.get("id")) not in previous_ids
        for e in repaired_evidence
    ):
        repaired_data["evidence"] = previous_data.get("evidence") or []

    return repaired_data


@cache_success_only(ttl=3600)
def cached_gemini_repair(repair_hash, repair_prompt):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": "validation_repair"}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=repair_prompt,
            config=config,
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="validation_repair",
            prompt_hash=repair_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        if not response.candidates:
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "msg": "Validation repair returned no candidates.", "debug": debug_info}
        candidate = response.candidates[0]
        finish_reason = str(candidate.finish_reason)
        debug_info["finish_reason"] = finish_reason
        if finish_reason == "FinishReason.MAX_TOKENS":
            debug_info["error_type"] = "MAX_TOKENS_REPAIR"
            return {"data": None, "error": True, "msg": "Validation repair reached the generation limit.", "debug": debug_info}
        if finish_reason and finish_reason != "FinishReason.STOP":
            debug_info["error_type"] = "Early Stop"
            return {"data": None, "error": True, "msg": f"Validation repair stopped early: {finish_reason}", "debug": debug_info}
        text = getattr(response, "text", None)
        if not text:
            try: text = response.candidates[0].content.parts[0].text
            except Exception: text = None
        if not text:
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "msg": "Validation repair returned empty text.", "debug": debug_info}
        try:
            data = extract_json(text)
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            return {"data": None, "error": True, "msg": "Validation repair produced invalid JSON.", "debug": debug_info}
        
        data = normalize_dossier_status_values(data)
        data = normalize_dossier_basis_values(data)
        data = sanitize_unsupported_not_currently_triggered_statuses(data)
        data = sanitize_unverifiable_verified_jurisdiction(data, str(address))
        try:
            _as_of = datetime.strptime(str(project_date), "%Y-%m-%d").date()
        except Exception:
            _as_of = date.today()
        data = sanitize_invalid_current_codes(data, _as_of)
        data = sanitize_invalid_evidence_propositions(data)
        data = sanitize_generic_proposition_links(data)
        data = sanitize_invalid_authority_evidence_links(data)
        data = sanitize_cross_discipline_applicability_links(data) # ADDED HERE
        data = repair_missing_threshold_links(data)
        data = sanitize_unsupported_permit_conclusions(data)
        data = sanitize_unsupported_threshold_conclusions(data)
        data = sanitize_unsupported_entitlement_rules(data)
        data = sanitize_unsupported_entitlement_conclusions(data)
        data = sanitize_unsupported_pathway_conclusions(data)
        data = sanitize_unverifiable_verified_pathways(data)
        data = sanitize_unverifiable_verified_permits(data)
        data = sanitize_semantically_misplaced_permit_findings(data)
        data = sanitize_unsubstantiated_conditional_statuses(data)
        data = sanitize_bottom_line_for_jurisdiction(data)
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_final_regulatory_statuses(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_validation_blocking_permit_findings(data)
        data = normalize_dossier_status_values(data)
        data = sanitize_bottom_line_for_unestablished_permits(data)
        data = sanitize_bottom_line_against_final_matrix(data)
        
        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        debug_info["validation_errors"] = validation_errors
        if validation_errors:
            debug_info["error_type"] = "Validation Failed After Repair"
            return {"data": data, "error": True, "msg": "Dossier still failed regulatory consistency validation after repair.", "debug": debug_info}
        debug_info["status"] = "success"
        return {"data": data, "error": False, "debug": debug_info}
    except Exception as e:
        debug_info["exception"] = str(e)
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="validation_repair",
            prompt_hash=repair_hash, elapsed_seconds=(time.monotonic() - _api_started) if "_api_started" in locals() else None,
            error=str(e)
        )
        debug_info["error_type"] = "Python Exception"
        return {"data": None, "error": True, "msg": f"Validation repair error: {str(e)[:200]}", "debug": debug_info}

# ============================================================
# DETERMINISTIC PERMIT DECISION ENGINE
# ============================================================
def derive_permit_statuses_from_evidence(data):
    """Make permit status a deterministic consequence of validated evidence.

    Gemini may discover evidence, but it cannot promote a discipline to a definitive
    permit status.  This function is the single source of truth for that decision.
    """
    if not isinstance(data, dict):
        return data

    evidence_by_id = {
        str(e.get("id")): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    for item in data.get("disciplines") or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown")
        ids = item.get("permit_evidence") or []
        if isinstance(ids, str):
            ids = [ids]
        if not isinstance(ids, list):
            ids = []

        valid_req = evidence_ids_supporting_type(
            ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline
        )
        valid_exempt = evidence_ids_supporting_type(
            ids, evidence_by_id, "PERMIT_EXEMPTION", discipline
        )

        if valid_exempt:
            item["permit"] = "NOT_APPLICABLE"
            item["permit_basis"] = "DIRECT_EVIDENCE"
            item["permit_evidence"] = valid_exempt
            item["permit_finding"] = (
                f"The available authoritative evidence establishes that a {discipline.lower()} "
                "permit is not required for the stated scope."
            )
        elif valid_req:
            item["permit"] = "VERIFIED_REQUIRED"
            item["permit_basis"] = "DIRECT_EVIDENCE"
            item["permit_evidence"] = valid_req
            item["permit_finding"] = (
                f"The available authoritative evidence establishes a {discipline.lower()} "
                "permit requirement for the stated scope."
            )
        else:
            # Preserve a genuinely unresolved status, but never preserve a Gemini
            # conclusion that implies a definitive permit consequence.
            current = str(item.get("permit") or "UNKNOWN").upper()
            if current in {"VERIFIED_REQUIRED", "REQUIRED", "PERMIT REQUIRED", "INFERRED", "NOT_APPLICABLE"}:
                item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"The available evidence does not yet establish whether a {discipline.lower()} "
                "permit is required. Check the applicable permit-specific authoritative source."
            )

    return data

# ============================================================
# FINAL DETERMINISTIC CONTRACT PASS
# ============================================================
def sanitize_jurisdiction_mismatched_process_evidence(data):
    """Remove process evidence that belongs to a different governmental jurisdiction."""
    if not isinstance(data, dict):
        return data
    jurisdiction = data.get("jurisdiction") or {}
    evidence = data.get("evidence") or []
    evidence_by_id = {str(e.get("id")): e for e in evidence if isinstance(e, dict) and e.get("id")}
    bad_ids = {
        eid for eid, ev in evidence_by_id.items()
        if ev.get("proposition_type") in {"PATHWAY", "REVIEW_REQUIREMENT", "AUTHORITY_HIERARCHY"}
        and _jurisdiction_process_source_mismatch(ev, jurisdiction)
    }
    if not bad_ids:
        return data

    for item in data.get("disciplines") or []:
        if not isinstance(item, dict):
            continue
        for key in ("process_evidence", "pathway_evidence", "review_evidence", "authority_evidence"):
            ids = item.get(key) or []
            if isinstance(ids, str):
                ids = [ids]
            if isinstance(ids, list):
                item[key] = [eid for eid in ids if str(eid) not in bad_ids]
        remaining = item.get("process_evidence") or []
        if not remaining:
            for key in ("process_reviewer", "process_route", "process_relationship", "process_trigger"):
                item.pop(key, None)

    bottom_ids = data.get("bottom_line_evidence") or []
    if isinstance(bottom_ids, str):
        bottom_ids = [bottom_ids]
    if isinstance(bottom_ids, list):
        data["bottom_line_evidence"] = [eid for eid in bottom_ids if str(eid) not in bad_ids]

    data["evidence"] = [e for e in evidence if str(e.get("id")) not in bad_ids]
    return data


def derive_process_status_from_evidence(data):
    """Derive a separate process summary from validated process evidence.

    PATHWAY evidence establishes routing/process. REVIEW_REQUIREMENT evidence establishes
    a documented review/inspection/submittal obligation but not necessarily the exact route.
    This does not alter the legacy pathway field.
    """
    if not isinstance(data, dict):
        return data
    data = sanitize_jurisdiction_mismatched_process_evidence(data)
    evidence_by_id = {
        str(e.get("id")): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    for item in data.get("disciplines") or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown").strip()
        ids = item.get("pathway_evidence") or []
        if isinstance(ids, str):
            ids = [ids]
        process_ids = item.get("process_evidence") or []
        if isinstance(process_ids, str):
            process_ids = [process_ids]
        combined = list(dict.fromkeys([str(x) for x in ids + process_ids]))
        valid_pathway = []
        valid_review = []
        for eid in combined:
            ev = evidence_by_id.get(eid)
            if evidence_supports_pathway(ev, discipline):
                valid_pathway.append(eid)
            elif evidence_supports_review(ev, discipline):
                valid_review.append(eid)
        if not valid_pathway and str(item.get("pathway") or "").upper() == "VERIFIED_REQUIRED":
            item["pathway"] = "CONDITIONAL"
            item["pathway_basis"] = "NOT_ESTABLISHED"
            item["pathway_evidence"] = []
            item["pathway_finding"] = "The available evidence does not yet establish the submission or review pathway."

        if valid_pathway:
            item["process_status"] = "PATHWAY_ESTABLISHED"
            item["process_basis"] = "DIRECT_EVIDENCE"
            item["process_finding"] = "A documented submission or review pathway is established by the cited source."
            item["process_evidence"] = list(dict.fromkeys(valid_pathway + valid_review))[:5]
            # PATHWAY evidence is strong enough to establish the legacy review/pathway
            # field too.  Keep the update deterministic so Gemini cannot set the status.
            item["pathway"] = "VERIFIED_REQUIRED"
            item["pathway_basis"] = "DIRECT_EVIDENCE"
            item["pathway_evidence"] = list(dict.fromkeys(valid_pathway))[:5]
            item["pathway_finding"] = "A documented submission or review pathway is established by the cited authoritative source."
        elif valid_review:
            item["process_status"] = "REVIEW_REQUIRED"
            item["process_basis"] = "DIRECT_EVIDENCE"
            item["process_finding"] = "A documented review or inspection requirement is established; exact routing remains unresolved."
            item["process_evidence"] = list(dict.fromkeys(valid_review))[:5]
        else:
            item["process_status"] = "NOT_ESTABLISHED"
            item["process_basis"] = "NOT_ESTABLISHED"
            item["process_finding"] = "The review/submission process has not yet been established by proposition-specific evidence."
            item["process_evidence"] = []
            for key in ("process_reviewer", "process_route", "process_relationship", "process_trigger"):
                item.pop(key, None)
    return data


def run_deterministic_contract_pass(data, address_text, project_date_value, selected_state=""):
    """Apply deterministic repairs once more, then run the full validators."""
    if not isinstance(data, dict):
        return data, ["Dossier output is not a JSON object."]
    data = normalize_dossier_status_values(data)
    data = sanitize_scope_relevance(data)
    data = sanitize_expected_process(data)
    data = sanitize_unverified_jurisdiction_identity(data)
    data = sanitize_jurisdiction_mismatched_process_evidence(data)
    data = derive_process_status_from_evidence(data)
    data = normalize_dossier_basis_values(data)
    data = sanitize_unsupported_not_currently_triggered_statuses(data)
    data = sanitize_unverifiable_verified_jurisdiction(data, str(address_text or ""))
    try:
        as_of = datetime.strptime(str(project_date_value), "%Y-%m-%d").date()
    except Exception:
        as_of = date.today()
    data = sanitize_invalid_current_codes(data, as_of)
    data = sanitize_invalid_evidence_propositions(data)
    data = derive_permit_statuses_from_evidence(data)
    data = sanitize_contradictory_established_pathways(data)
    data = sanitize_generic_proposition_links(data)
    data = sanitize_invalid_authority_evidence_links(data)
    data = sanitize_cross_discipline_applicability_links(data)
    data = repair_missing_threshold_links(data)
    data = sanitize_unsupported_permit_conclusions(data)
    data = sanitize_unsupported_threshold_conclusions(data)
    data = sanitize_unsupported_entitlement_rules(data)
    data = sanitize_unsupported_entitlement_conclusions(data)
    data = sanitize_unsupported_pathway_conclusions(data)
    data = sanitize_unverifiable_verified_pathways(data)
    data = sanitize_unverifiable_verified_permits(data)
    data = sanitize_semantically_misplaced_permit_findings(data)
    data = sanitize_unsubstantiated_conditional_statuses(data)
    data = sanitize_bottom_line_for_jurisdiction(data)
    data = sanitize_bottom_line_for_unestablished_permits(data)
    data = sanitize_bottom_line_against_final_matrix(data)
    data = normalize_dossier_status_values(data)
    data = sanitize_final_regulatory_statuses(data)
    data = normalize_dossier_status_values(data)
    data = sanitize_validation_blocking_permit_findings(data)
    data = normalize_dossier_status_values(data)
    data = sanitize_bottom_line_for_unestablished_permits(data)
    data = sanitize_bottom_line_against_final_matrix(data)
    data = derive_permit_statuses_from_evidence(data)
    data = sanitize_bottom_line_for_unestablished_permits(data)
    data = sanitize_bottom_line_against_final_matrix(data)
    data = derive_permit_statuses_from_evidence(data)
    data = sanitize_bottom_line_for_unestablished_permits(data)
    data = sanitize_bottom_line_against_final_matrix(data)
    data = derive_process_status_from_evidence(data)
    data = sanitize_bottom_line_against_final_matrix(data)
    errors = validate_dossier(data)
    errors.extend(validate_bottom_line(data))
    errors.extend(validate_input_state_consistency(data, address_text, selected_state))
    return data, errors

def sanitize_scope_relevance(data):
    """Keep scope classifications useful without letting them become permit conclusions."""
    if not isinstance(data, dict):
        return data
    allowed = {"LIKELY", "PROBABLE", "POSSIBLE", "UNLIKELY"}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        value = str(item.get("scope_likelihood") or "").upper().strip()
        if value not in allowed:
            value = "POSSIBLE"
        item["scope_likelihood"] = value
        basis = str(item.get("scope_basis") or "").strip()
        item["scope_basis"] = basis[:400]
    return data

def sanitize_expected_process(data):
    """Process expectations must remain source-grounded and must not assert permit requirements."""
    if not isinstance(data, dict):
        return data
    allowed = {"ESTABLISHED", "EXPECTED", "NOT_ESTABLISHED"}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        conf = str(item.get("process_confidence") or "").upper().strip()
        if conf not in allowed:
            conf = "NOT_ESTABLISHED"
        process = str(item.get("expected_process") or "").strip()
        if not process:
            conf = "NOT_ESTABLISHED"
        # Do not allow an expected-process field to smuggle in a definitive permit claim.
        if re.search(r"\b(?:permit\s+(?:is\s+)?required|requires?\s+(?:a\s+)?permit|must\s+obtain\s+(?:a\s+)?permit)\b", process, re.I):
            process = ""
            conf = "NOT_ESTABLISHED"
        item["process_confidence"] = conf
        item["expected_process"] = process[:600]
    return data

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state: st.session_state.error_msg = None
if "run_status" not in st.session_state: st.session_state.run_status = "Ready"
if "run_started_utc" not in st.session_state: st.session_state.run_started_utc = None
if "run_id" not in st.session_state: st.session_state.run_id = None

st.title("🏛️ AHJ Research Assistant v27.00")
st.caption("Research brief: jurisdiction + current codes + scope-based disciplines + targeted questions + expected AHJ process. Strict evidence for regulatory conclusions; clearly labeled research leads for unresolved issues.")

with st.sidebar:
    st.warning("⚠️ Pay-As-You-Go Active. Successful research results cached for 1 hour; failed calls are never cached.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=False)

st.header("1. Project Metadata")
col1, col2 = st.columns(2)
with col1:
    state = st.selectbox("State / Jurisdiction", STATE_OPTIONS, index=STATE_OPTIONS.index("Oregon"))
    address = st.text_input("Project Address", "24340 NW Meek Rd, 97124")
    project_date = st.date_input("Permit / Construction Date", date.today())
with col2:
    ptype = st.selectbox("Project Type", PROJECT_TYPES, index=0)
    bclass = st.selectbox("Building / Occupancy Class", BUILDING_CLASSES, index=0)
    existing_permit = st.text_input("Existing Entitlements (Optional)", "", placeholder="e.g., CUP, variance, site plan")

st.header("2. Scope of Work (SOW)")
sow_text = st.text_area("Paste the complete Scope of Work below.", height=200,
    value="""Ground-level exterior commercial HVAC unit replacement.
Replacement is intended to be like-for-like but will be a different brand.
The unit will remain in the same location on the existing exterior pad.
Existing ductwork will not be modified.
No roof penetrations are proposed.
The parcel operates under an existing Conditional Use Permit.
Electrical scope, replacement equipment specifications, equipment weight,
anchorage details, and existing CUP conditions have not yet been verified.""")

st.header("3. Research Execution")
input_string = f"{PROMPT_VERSION}|{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    st.session_state.error_msg = None
    st.session_state.report_data = None
    st.session_state.run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
    st.session_state.run_started_utc = datetime.now(timezone.utc).isoformat()
    st.session_state.run_status = "Started — validating project inputs"
    if mock_mode:
        st.session_state.report_data = {
            "bottom_line": "Mechanical permit requirement is established. Specific review pathway remains conditional. Electrical scope is unknown. Structural and Planning determinations require retrieval of governing conditions and equipment specifications.",
            "bottom_line_evidence": ["E2", "E3", "E4"],
            "research_completeness": {"status": "PARTIAL", "reason": "Main regulatory framework established, but equipment specs and CUP conditions remain unresolved.", "critical_missing": ["Equipment cut sheets", "Existing CUP document"]},
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Unincorporated", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Jurisdiction", "proposition_type": "JURISDICTION", "source_type": "permit_page", "retrieval_note": "Confirms jurisdiction for unincorporated areas.", "rule": "Jurisdiction for unincorporated areas requires county confirmation."},
                {"id": "E2", "title": "Washington County Mechanical Permit Requirements", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Mechanical", "proposition_type": "PERMIT_REQUIREMENT", "source_type": "permit_page", "retrieval_note": "County permit requirements for commercial mechanical work.", "rule": "Commercial HVAC equipment replacement requires a mechanical permit."},
                {"id": "E3", "title": "Oregon Energy Efficiency Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Energy", "proposition_type": "APPLICABILITY", "source_type": "code", "retrieval_note": "Governs replacement equipment efficiency.", "rule": "Replacement mechanical equipment must comply with current energy efficiency standards."},
                {"id": "E4", "title": "Oregon Electrical Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Electrical", "proposition_type": "PERMIT_REQUIREMENT", "source_type": "code", "retrieval_note": "Defines when electrical modifications trigger permits.", "rule": "Electrical permit required for modification of branch circuits or disconnects."}
            ],
            "disciplines": [
                {"type": "Mechanical", "scope_likelihood": "LIKELY", "scope_basis": "The SOW directly replaces commercial HVAC equipment.", "applicability": {"rule": "Commercial mechanical equipment replacement requires a mechanical permit.", "fact": {"statement": "Replacing ground-level exterior commercial HVAC equipment.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E2"]}, "permit": "VERIFIED_REQUIRED", "permit_finding": "Mechanical permit requirement is established by the cited permit/code evidence.", "permit_basis": "DIRECT_EVIDENCE", "permit_evidence": ["E2"], "pathway": "CONDITIONAL", "pathway_finding": "Specific review pathway remains unresolved because project-specific pathway criteria have not been verified.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Project-specific review pathway criteria (e.g., unit weight, CFM)."], "reopen": ["Authoritative pathway criteria retrieved."]},
                {"type": "Electrical", "scope_likelihood": "PROBABLE", "scope_basis": "The SOW identifies electrical scope as unresolved for replacement equipment.", "applicability": {"rule": "Electrical permit required for branch circuit/disconnect modification.", "fact": {"statement": "Electrical modifications to replacement unit are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Whether wiring/disconnect/breaker will be modified.", "relationship": "conditional", "evidence": ["E4"]}, "permit": "CONDITIONAL", "permit_finding": "Electrical permit consequence depends on whether wiring/disconnect/circuit work occurs.", "permit_basis": "CONDITIONAL", "permit_evidence": ["E4"], "pathway": "CONDITIONAL", "pathway_finding": "Pathway cannot be determined until electrical scope is defined.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Unit electrical specs (MCA, MOP, voltage).", "Scope of electrical changes."], "reopen": ["Modifying electrical disconnect, wiring, or breaker."]},
                {"type": "Energy", "scope_likelihood": "LIKELY", "scope_basis": "HVAC replacement can implicate current energy-code requirements.", "applicability": {"rule": "Current energy code applies to replacement equipment.", "fact": {"statement": "Replacing HVAC unit; efficiency ratings unknown.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E3"]}, "permit": "CONDITIONAL", "permit_finding": "Energy compliance applies; a separate energy permit requirement is not established by current evidence.", "permit_basis": "NOT_ESTABLISHED", "permit_evidence": [], "pathway": "CONDITIONAL", "pathway_finding": "Compliance pathway depends on equipment specifications and primary permit type.", "pathway_basis": "NOT_ESTABLISHED", "pathway_evidence": [], "missing": ["Replacement equipment efficiency/specifications."], "reopen": []},
                {"type": "Structural", "scope_likelihood": "POSSIBLE", "scope_basis": "Equipment weight and anchorage are unresolved.", "applicability": {"rule": "Structural requirements depend on replacement equipment loads and attachment conditions.", "fact": {"statement": "Replacement unit weight and anchorage configuration are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Replacement unit weight and anchorage details.", "relationship": "conditional", "evidence": []}, "permit": "UNKNOWN", "permit_finding": "Structural permit consequence cannot be determined from current project facts.", "permit_basis": "NOT_ESTABLISHED", "permit_evidence": [], "pathway": "UNKNOWN", "pathway_finding": "Structural review pathway cannot be established from current information.", "pathway_basis": "NOT_ESTABLISHED", "pathway_evidence": [], "missing": ["Replacement unit operating weight", "Anchorage configuration"], "reopen": ["Equipment weight or anchorage requires structural review"]},
                {"type": "Planning / CUP", "scope_likelihood": "POSSIBLE", "scope_basis": "An existing CUP is identified and exterior equipment work is proposed.", "applicability": {"rule": "Work must comply with existing CUP conditions.", "fact": {"statement": "Parcel operates under existing CUP; actual conditions not retrieved.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Actual CUP conditions governing exterior equipment.", "relationship": "conditional", "evidence": ["E1"]}, "permit": "CONDITIONAL", "permit_finding": "Existing CUP identified, but governing conditions have not been reviewed.", "permit_basis": "CONDITIONAL", "permit_evidence": [], "pathway": "CONDITIONAL", "pathway_finding": "Final land-use determination conditional on review of existing CUP conditions.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Actual CUP conditions governing exterior equipment."], "reopen": ["Relocation, footprint expansion, screening changes, noise increases, or site work occur."]}
            ]
        }
        st.session_state.debug_log = {"mock": True, "note": "No API call made", "api_usage": None}
        st.success("🛡️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.session_state.error_msg = "GEMINI_KEY missing."
        else:
            with st.spinner("Performing deep authoritative research..."):
                prompt = f"""
You are an expert AHJ research analyst. Return ONLY valid JSON. No conversational text, markdown, or explanations outside the JSON.
PROJECT:
State: {state} | Address: {address} | Date: {project_date}
Type: {ptype} | Class: {bclass} (USER-PROVIDED) | Entitlements: {existing_permit}
SCOPE: {sow_text}
RESEARCH CONTRACT:
Research deeply, but write compactly. The SOW may be short or long. Never assume missing facts.
CODE CURRENCY FIREWALL — CRITICAL:
The project date is the as-of date for code currency. Do NOT use remembered code editions.
JURISDICTION FIREWALL: Determine the project's actual governmental jurisdiction from authoritative site-specific evidence (parcel/GIS/property record/jurisdiction lookup or an equivalent official source). A postal city, ZIP code, mailing address city, or generic county/city service page does NOT establish municipal jurisdiction. Merely repeating the street address on a municipal page does not establish that the parcel is inside municipal limits. If the parcel is unincorporated, set the actual city field to "Unincorporated" and do not treat the postal city as the municipal jurisdiction. The AHJ must correspond to the verified governmental jurisdiction.
AUTHORITY-HIERARCHY / LOCAL-OVERRIDE FIREWALL — CRITICAL:
Do NOT assume that a state code automatically controls the project. First research the legal relationship between state and local authority for the relevant discipline in the verified jurisdiction. Some states have statewide mandatory codes with limited local amendments; some delegate enforcement to local governments; some permit local amendments; some home-rule jurisdictions can adopt provisions that modify or exceed state baselines. This is a research question, not a model assumption.
For each material permit conclusion, determine: (1) what state rule says, (2) whether the state rule controls in this jurisdiction, (3) whether the local AHJ has adopted amendments or independent requirements, and (4) which rule is controlling for THIS project.
If a state-level PERMIT_REQUIREMENT is used to establish VERIFIED_REQUIRED, also retrieve AUTHORITY_HIERARCHY evidence showing why the state rule controls or remains applicable in this jurisdiction. If local law controls or may modify the state rule, use the local controlling rule instead. If the authority relationship cannot be established, do not present the state rule as unqualified VERIFIED_REQUIRED; use CONDITIONAL/UNKNOWN until the hierarchy is resolved.
Do not use the phrase "home rule" as a shortcut. Verify the actual statutory/code framework and the actual local adoption/amendment status.
For every code listed as CURRENT, actively research the jurisdiction's official code-adoption/current-code source using Google Search.
Prefer the official state, county, city, or AHJ adoption page over secondary summaries.
The evidence proposition MUST be CODE_CURRENCY and its rule MUST establish the adopted edition and/or effective/mandatory/current status.
Do not mark an older code CURRENT merely because it remains available online or is a prior code edition.
If a newer edition has become mandatory by the project date, the older edition is not CURRENT.
If a phase-in period legally permits either edition, describe that explicitly and do not silently label the older edition CURRENT unless the authoritative source says it is current/permitted for the project date.
If code currency cannot be established from authoritative evidence, use status = CONDITIONAL and identify the code-currency evidence gap.
Code name/year MUST agree with the cited CODE_CURRENCY evidence.
Never use a generic applicability source as proof of code currency.
DISCIPLINE DISCOVERY: Determine disciplines dynamically from the project scope, project type, jurisdiction, adopted codes, land-use controls, and authoritative applicability rules. Do not use a fixed discipline list. Do not impose a maximum number of disciplines.

SCOPE-BASED RESEARCH PRIORITY — CRITICAL:
The primary purpose of this tool is to help a human start regulatory research on the right path. For every discipline that is plausibly implicated by the SOW, provide a scope_likelihood using ONLY these values: LIKELY, PROBABLE, POSSIBLE, UNLIKELY. This is a scope/research classification, NOT a permit conclusion. Base it on explicit SOW facts, project metadata, and authoritative applicability information when available. Never use a permit assumption to make a discipline LIKELY.
- LIKELY = the SOW directly describes work normally belonging to this discipline and the project facts clearly implicate it.
- PROBABLE = the SOW strongly suggests the discipline is involved, but an important scope fact remains unresolved.
- POSSIBLE = a concrete scope item creates a legitimate regulatory question, but the discipline may not actually be involved.
- UNLIKELY = there is no meaningful scope signal; normally omit such a discipline unless an authoritative rule makes it relevant.
Each discipline must also include scope_basis: one short explanation tied to the actual SOW. Do not invent scope facts.

RESEARCH-BRIEF OBJECTIVE — CRITICAL:
Do NOT optimize the dossier around proving a permit requirement. The useful output is a research map: (1) actual jurisdiction, (2) current governing codes, (3) disciplines likely/probable/possible from the SOW, (4) what is known about applicability, permit consequences, and process, (5) concrete questions a human can take to the AHJ, and (6) clearly labeled possible research leads. A permit conclusion remains strict and evidence-based, but an unresolved permit question is still a successful research result when the dossier gives the human a useful next step.

EXPECTED PROCESS — CRITICAL:
For each material discipline, provide expected_process only when it is supported by an authoritative AHJ/state/local source. It may describe the normal submission/review route for that discipline even when the project-specific permit consequence remains unresolved, but label process_confidence as ESTABLISHED when the source directly applies to the discipline/obligation and EXPECTED when it is a general AHJ process for that type of work. Never present a generic portal or agency page as proof that THIS project requires a permit. If no useful process source is found, use process_confidence = NOT_ESTABLISHED and leave expected_process blank.
For every discipline determine separately:
1. APPLICABILITY: Does the authoritative rule apply to the known project facts?
2. PERMIT: Does authoritative evidence establish a permit requirement or exemption?
3. PATHWAY: Does authoritative evidence establish a specific review pathway?
4. AUTHORITY HIERARCHY: When state and local rules could differ, which governmental rule controls this subject in this jurisdiction?

TARGETED PERMIT RESEARCH — CRITICAL:
When a material permit question remains unresolved, make a focused attempt to find the legal permit consequence during this research pass, but do not let the permit question consume the entire research budget. Search the verified AHJ's authoritative sources using the actual work item plus permit language. If no proposition-specific permit rule is found, preserve the unresolved status and spend the remaining research effort on useful process information and targeted human questions.
Prioritize official permit requirement pages/checklists, official code or ordinance provisions, official applications/fee schedules that expressly identify the permit, and official AHJ FAQs/interpretations. Generic homepages, portals, and code indexes are not permit evidence unless the page itself states the proposition.
Never manufacture a permit rule from common practice, a page title, URL, section number, regulated-work language, or application existence.
REGULATORY FIREWALL:
SOURCE RULE + PROJECT FACT does NOT automatically prove a permit or pathway.
Never convert: threshold → permit; threshold → exemption; exemption → pathway; code applicability → permit; permit → plan-review pathway; existing entitlement → exemption.
Permit = VERIFIED_REQUIRED only with permit-specific evidence AND, when the permit evidence is state-level, separate authority-hierarchy evidence establishing that the state rule controls or remains applicable in this jurisdiction.
Pathway = VERIFIED_REQUIRED only with pathway-specific evidence.
Use CONDITIONAL when unresolved facts could change the result. Use UNKNOWN when evidence is insufficient. Use NOT_APPLICABLE only when authoritative evidence establishes non-applicability for the known project facts.
INTERNAL CONSISTENCY RULES:
The following combinations are invalid:
A. determination = does_not_apply AND missing is non-empty
B. determination = cannot_determine AND permit = NOT_APPLICABLE
C. determination = cannot_determine AND pathway = NOT_APPLICABLE
D. permit = NOT_APPLICABLE AND applicability determination != does_not_apply
E. pathway = NOT_APPLICABLE AND applicability determination != does_not_apply
F. determination = does_not_apply AND relationship != direct
G. permit = VERIFIED_REQUIRED AND permit_basis != DIRECT_EVIDENCE
H. pathway = VERIFIED_REQUIRED AND pathway_basis != DIRECT_EVIDENCE
I. permit_basis = DIRECT_EVIDENCE AND permit_evidence is empty
J. pathway_basis = DIRECT_EVIDENCE AND pathway_evidence is empty
If applicability cannot be determined because a material fact is unknown, use: determination = cannot_determine, relationship = conditional, permit = CONDITIONAL or UNKNOWN, pathway = CONDITIONAL or UNKNOWN.
Never use NOT_APPLICABLE as a placeholder for "I don't know." Never use does_not_apply as a placeholder for "missing information."
DISCIPLINE EVIDENCE: Applicability evidence must support the applicability rule for that discipline. Do not use Mechanical evidence to establish Planning/Zoning/CUP applicability. Do not use Electrical evidence to establish Structural applicability. If the correct discipline-specific source cannot be found, use cannot_determine / UNKNOWN rather than borrowing unrelated evidence.
ENERGY: Do not confuse applicability with compliance data. Missing equipment efficiency ratings do not by themselves make energy-code applicability UNKNOWN if the authoritative rule already establishes that the replacement work is within the energy-code scope. If the energy rule applies but compliance details are unknown: applicability = applies, permit = CONDITIONAL or UNKNOWN, pathway = CONDITIONAL or UNKNOWN.
PROJECT FACT INTEGRITY — CRITICAL:
A project fact is USER_PROVIDED only when explicitly stated in the SOW or project metadata.
A project fact is RETRIEVED_RECORD only when established by a retrieved project-specific record.
A project fact is AUTHORITATIVE_SOURCE only when the authoritative source itself establishes the fact.
INFERRED means the model logically suspects something may be true, but it is NOT an established project fact.
UNKNOWN means the fact is not established.
NEVER convert a normal construction assumption into a project fact.
For electrical work, NEVER assume: electrical reconnection, existing circuit reuse, disconnect reuse, disconnect replacement, breaker replacement, conductor replacement, wiring modification, MCA/MOP, voltage, phase, circuit capacity, equipment connection method.
For structural work, NEVER assume: replacement unit weight, operating weight, anchorage, attachment method, structural capacity, framing condition, roof loading.
For planning/land use, NEVER assume: CUP conditions, zoning compliance, setbacks, screening compliance, noise compliance, site-plan compliance, entitlement conditions.
IMPORTANT:
An INFERRED fact may identify something worth investigating, but it cannot establish a DIRECT applicability relationship.
If applicability depends on an inferred or unknown project fact:
determination = cannot_determine
relationship = conditional
and identify the missing fact.
Do not use INFERRED facts to produce: direct applicability, does_not_apply, VERIFIED_REQUIRED permit, or VERIFIED_REQUIRED pathway.
ELECTRICAL EXAMPLE:
If the SOW says electrical scope is unknown, do NOT infer that the replacement unit will be electrically reconnected.
Correct:
{{
"fact": {{
"statement": "Electrical modification scope is unknown.",
"source": "USER_PROVIDED"
}},
"determination": "cannot_determine",
"relationship": "conditional"
}}
Incorrect:
{{
"fact": {{
"statement": "Replacement unit will require electrical reconnection.",
"source": "INFERRED"
}},
"determination": "applies",
"relationship": "direct"
}}
The second pattern is prohibited even if electrical reconnection would be common construction practice. Common construction practice ≠ established project fact.
EXISTING ENTITLEMENTS: If a CUP, variance, site plan, development agreement, or other entitlement is identified, do not assume its conditions. Retrieve the governing document when it could affect the result. Do not conclude that an amendment is unnecessary without supporting authoritative evidence.
SUBSTANTIVE EVIDENCE REQUIREMENT — CRITICAL:
An official source being in the correct discipline does NOT by itself prove a permit requirement or review pathway.
For every permit conclusion, distinguish:
1. CODE APPLICABILITY: Evidence that a code, ordinance, or regulation applies to the project.
2. PERMIT REQUIREMENT: Evidence that a permit or approval is actually required.
3. REVIEW PATHWAY: Evidence establishing how that permit or approval is processed (plan review, trade permit, over-the-counter, engineering review, etc.).
IMPORTANT:
"Code applies" does NOT mean "permit required."
"Compliance is required" does NOT mean "permit required."
"Subject to the mechanical code" does NOT mean "mechanical permit required."
"Equipment exceeds a threshold" does NOT by itself establish a permit, plan review, engineering review, or other downstream regulatory consequence.
A VERIFIED_REQUIRED permit conclusion is allowed only when an authoritative source directly establishes the permit requirement for the applicable project activity.
A state-level permit rule is NOT enough by itself where local authority may modify, supersede, or independently regulate that subject. Resolve the state/local authority hierarchy first.
A VERIFIED_REQUIRED pathway conclusion is allowed only when an authoritative source directly establishes that pathway.
If the source establishes only code applicability, keep the permit conclusion CONDITIONAL or UNKNOWN unless separate permit evidence is found.
If a conditional permit finding says a condition "requires" or "triggers" a permit, separate PERMIT_REQUIREMENT evidence is still mandatory. Otherwise state that the permit consequence is not established.
If permit evidence exists but pathway evidence does not, the correct result is: permit = VERIFIED_REQUIRED, permit_basis = DIRECT_EVIDENCE, pathway = CONDITIONAL or UNKNOWN, pathway_basis = NOT_ESTABLISHED.
BASIS VOCABULARY IS CLOSED — CRITICAL:
IMPORTANT STATUS RULE: NOT_CURRENTLY_TRIGGERED is NOT a synonym for NOT_ESTABLISHED. Use NOT_CURRENTLY_TRIGGERED only when the current applicability is established and authoritative evidence supports that the obligation is presently untriggered. If evidence is insufficient, use CONDITIONAL or UNKNOWN with basis NOT_ESTABLISHED. Never write a definitive "no permit is required" or "no separate permit is issued" statement without valid PERMIT_EXEMPTION evidence.
For permit_basis and pathway_basis, use ONLY:
- DIRECT_EVIDENCE
- CONDITIONAL
- NOT_ESTABLISHED
Never emit INSUFFICIENT_EVIDENCE, INSUFFICIENT, UNSUPPORTED, NOT_SUPPORTED, or other synonyms.
If evidence is missing or insufficient, use NOT_ESTABLISHED. Do not invent a new basis value.
Never infer a permit requirement from REVIEW_REQUIREMENT, PATHWAY, THRESHOLD, or APPLICABILITY evidence alone.
EVIDENCE PROPOSITION TYPES — CRITICAL:
Every evidence record MUST identify exactly what regulatory proposition the source establishes.
Use:
- JURISDICTION: establishes which AHJ has authority.
- AUTHORITY_HIERARCHY: establishes the legal relationship between state and local regulatory authority for the relevant subject (for example statewide control, delegated local enforcement, local amendment authority, home-rule authority, preemption, minimum-statewide standard, or local rule prevailing). This is separate from the permit requirement itself.
- CODE_CURRENCY: establishes adopted code edition or effective date.
- APPLICABILITY: establishes that a rule/code applies to the project activity.
- PERMIT_REQUIREMENT: explicitly establishes that a permit or approval is required.
- PERMIT_EXEMPTION: explicitly establishes that a permit or approval is not required or an exemption applies.
- REVIEW_REQUIREMENT: establishes review, inspection, engineering, calculations, or submittal requirements WITHOUT establishing that a permit is required.
- PATHWAY: establishes how an already-established permit/approval/review is submitted or processed.
- THRESHOLD: establishes a numeric or categorical limit/trigger. A threshold alone is NEVER a permit requirement.
EVIDENCE LABEL INTEGRITY — CRITICAL:
Choose proposition_type from the exact proposition stated by the source, not from the conclusion you want. The `rule` field must contain only what the cited source establishes.
If the source says "review is required," classify it as REVIEW_REQUIREMENT. If it says equipment over a weight threshold requires review, classify it as THRESHOLD and/or REVIEW_REQUIREMENT unless it separately states that a permit is required.
Do NOT relabel review, threshold, applicability, compliance, or pathway evidence as PERMIT_REQUIREMENT to preserve a permit conclusion. If explicit permit evidence cannot be found, downgrade the conclusion and identify the evidence gap.
- PERMIT_EXEMPTION: explicitly establishes that a permit is not required or an exemption applies.
- PATHWAY: explicitly establishes how an approval is processed (plan review, trade permit, over-the-counter, engineering review, application type, etc.).
- THRESHOLD: establishes a numerical or categorical threshold.
- ENTITLEMENT: establishes an actual project-specific CUP, variance, site plan, development agreement, or similar governing condition.
- OTHER: authoritative information that does not fit the categories above.
CRITICAL:
Do not label an evidence item PERMIT_REQUIREMENT merely because it discusses a code, compliance, or regulated work. The source itself must establish the permit requirement.
Do not label an evidence item PATHWAY merely because it is a permit page. The source must establish the actual review/process pathway.
Do not label an evidence item PERMIT_EXEMPTION unless the source explicitly establishes the exemption or non-requirement.
Each rule field must state the proposition actually supported by the source.
REVIEW REQUIREMENT: Use REVIEW_REQUIREMENT for engineering review, plan review, inspection, administrative review, or similar obligations when the source does not itself establish a permit requirement. Review requirement evidence MUST NOT be treated as PERMIT_REQUIREMENT evidence.
PATHWAY: Use PATHWAY for the processing route of an already-established obligation. A pathway source MUST NOT be used to prove that the underlying permit or approval is required.
PERMIT EXEMPTION: Use PERMIT_EXEMPTION only when the source explicitly establishes that the permit/approval is not required or an exemption applies. Do not infer an exemption from like-for-like work, replacement, repair, existing conditions, or common practice.
LEVEL-3 LOGICAL CHAIN FIREWALL — CRITICAL:
For planning/land-use work, an existing CUP or generic zoning applicability rule does not establish that a CUP amendment, land-use approval, or planning clearance is required or not required; require ENTITLEMENT evidence for that legal consequence.
A PERMIT_REQUIREMENT proposition must explicitly establish the permit/approval consequence. An APPLICABILITY, REVIEW_REQUIREMENT, PATHWAY, THRESHOLD, or CODE_CURRENCY proposition cannot be promoted into a permit conclusion merely because it appears relevant. Likewise, PATHWAY evidence cannot be promoted into a permit requirement. If the source does not state the downstream consequence, keep the consequence CONDITIONAL/UNKNOWN/NOT_ESTABLISHED.
THRESHOLD DETECTION — CRITICAL:
Do not describe ordinary equipment specifications, code edition years, section numbers, MCA/MOP labels, license numbers, or unrelated numeric values as regulatory thresholds.
For electrical work, MCA/MOP/ampacity/circuit terminology is a threshold proposition only when the source itself states a numeric or categorical limit tied to that terminology.
A project fact such as “MCA/MOP details are missing” is NOT itself a threshold claim.
THRESHOLD → CONSEQUENCE FIREWALL — CRITICAL:
A numerical threshold found in an authoritative source establishes only the proposition actually stated by that source.
Do NOT infer a permit requirement, engineering requirement, plan review, anchorage requirement, exemption, or pathway from a threshold unless the source explicitly establishes that consequence.
Always preserve the exact relationship: SOURCE RULE → PROJECT FACT → APPLICABILITY → EXPLICIT CONSEQUENCE.
BOTTOM LINE EVIDENCE FIREWALL — CRITICAL:
The Bottom Line may summarize conclusions already established in the discipline sections.
The Bottom Line MUST NOT introduce: a new permit requirement, a new permit type, a new threshold, a new exemption, a new review pathway, a new jurisdiction conclusion, a new CUP conclusion, or a new code applicability conclusion.
Every material Bottom Line conclusion must be traceable to one or more bottom_line_evidence IDs.
DECISION TRAIL — CRITICAL: For every VERIFIED_REQUIRED, NOT_APPLICABLE, or otherwise materially established conclusion, preserve the specific evidence IDs that support that proposition. The report UI will expose these as the human's audit trail. Never rely on a generic AHJ homepage when a proposition-specific permit, exemption, review, pathway, or code source was found.
For unresolved findings, applicability evidence may be shown as a research starting point, but it must never be described as proof of the unresolved permit/pathway consequence.
RESEARCH COMPLETENESS: SUFFICIENT = material conclusions supported by adequate authoritative evidence. PARTIAL = main framework established but material facts/documents remain unresolved. INSUFFICIENT = jurisdiction, governing code, permit authority, or material requirements cannot be established.
AI RESEARCH LEADS — CRITICAL:
A regulatory dossier can contain useful professional research leads even when the law/permit conclusion is not established. In the user-facing report these are labeled “Worth checking — AI-generated possible lead.” These are NOT regulatory conclusions and MUST NOT affect permit, pathway, applicability, jurisdiction, code status, Bottom Line, or research completeness.
For each discipline, optionally return 0-5 potential_issues only when the SOW and research suggest a concrete issue worth investigating. Zero is correct when no useful lead exists.
Each lead must be visibly framed as a possibility using language such as "may warrant", "could depend on", "worth checking", or "may require further review". Never state a model-inference lead as a fact, requirement, exemption, trigger, or definitive agency action.
A lead should be grounded in a specific SOW item, project fact, missing document, governing topic, or unresolved relationship. Do not invent risks merely because they are common in construction.
Good examples: ground-mounted HVAC equipment may warrant structural review depending on equipment weight/anchorage; accessibility alterations may warrant additional review depending on extent of alteration; an existing assembly use may warrant checking project-specific CUP/site-plan conditions.
Do not turn a lead into an actionable question automatically. A question should exist only when it resolves a concrete uncertainty.
TARGETED VERIFICATION QUESTIONS — CRITICAL:
When a discipline has permit = UNKNOWN/CONDITIONAL or pathway = UNKNOWN/CONDITIONAL, generate 0-5 specific, actionable questions for the user to research or ask the AHJ.
Do NOT generate generic questions such as "Is a permit required?", "Are there exemptions?", or "Where do I apply?".
A question is only actionable if the user could take it to the AHJ or use it in targeted research and it would resolve a particular uncertainty in THIS project.
Each question MUST synthesize at least one specific project fact from the SOW with the actual AHJ, governing code, entitlement, missing fact, or unresolved regulatory issue.
Never use a question merely to restate the status. For example, do not ask "what code determines whether this needs a permit?" when the dossier already says the permit is unknown.
If the research found no specific regulatory decision point, do NOT invent a question just to fill the array; return an empty actionable_questions array.
Prefer questions that name the actual work item, equipment, location, document, code section/topic, permit type, entitlement, or AHJ division involved.
The questions should help resolve the specific uncertainty shown in the discipline finding.
A question must have a concrete regulatory decision point behind it; do not ask for information merely because it is commonly useful.
If a missing fact would affect only engineering design or compliance, but not the permit/pathway determination, do not present it as a permit question.
For every VERIFIED_REQUIRED permit or VERIFIED_REQUIRED pathway, ensure the cited evidence is proposition-specific and the conclusion matches what that evidence actually establishes.
SOURCE-SPECIFICITY — CRITICAL: When a source supports a specific permit, exemption, pathway, code section, or regulatory proposition, cite the actual authoritative page, ordinance, code text, checklist, application instruction, or document that contains that proposition. Do NOT cite a generic agency homepage, department landing page, or broad code index merely because it belongs to the correct agency. A generic landing page is acceptable only when that page itself contains the proposition being claimed. Evidence titles must describe the actual source retrieved, not a guessed section title attached to a generic URL.
SOURCE TRACEABILITY — CRITICAL: A human reviewer must be able to open the cited URL and find the claimed proposition without relying on model knowledge. If the exact proposition-specific source cannot be located, do not manufacture a specific citation; keep the conclusion UNKNOWN/CONDITIONAL and identify the evidence gap.
Prefer questions that identify the exact regulatory decision point, such as whether a particular scope item triggers a separate permit, whether an existing approval governs the work, whether a stated code provision applies to the specific alteration, or what fact/document the AHJ needs to make the determination.
If the discipline has a missing project fact, turn that fact into a concrete verification question tied to the scope rather than merely repeating the missing-information label.
If a permit is already VERIFIED_REQUIRED but the pathway is unresolved, ask targeted questions about the actual submission/review route for that established permit rather than asking whether a permit is required.
Questions must be concise, practical, and written for a project manager/owner to use with the AHJ. There is NO minimum question count: five excellent questions are better than five filler questions, and zero is correct when no actionable decision point remains.
Store these questions in the dedicated actionable_questions array. Do not put generic questions there merely to fill the array.
VALIDATION-AWARE RESEARCH: Treat the evidence taxonomy as an enforcement contract. If authoritative permit evidence cannot be found, do not manufacture it by relabeling applicability, threshold, review, or pathway evidence. Prefer a precise CONDITIONAL/UNKNOWN result with an evidence gap. If an existing entitlement is identified but its governing document is unavailable, do not infer its conditions or amendment consequences.
OUTPUT: Keep JSON concise. Rule: 10-25 words. Fact: 5-15 words. Finding: 10-25 words. Missing/reopen item: short phrase. Research quality is more important than brevity.
JSON SCHEMA:
{{
"bottom_line": "3 concise sentences maximum",
"bottom_line_evidence": ["E1"],
"research_completeness": {{"status": "SUFFICIENT|PARTIAL|INSUFFICIENT", "reason": "short", "critical_missing": ["short item"]}},
"jurisdiction": {{"status": "VERIFIED|CONDITIONAL", "county": "string", "city": "string", "ahj": "string", "evidence": ["E1"]}},
"codes": [{{"name": "string", "status": "CURRENT|CONDITIONAL", "evidence": ["E2"]}}],
"evidence": [{{"id": "E1", "title": "string", "url": "string", "authority": "state|county|city|federal|tribal|other", "discipline": "string", "proposition_type": "JURISDICTION|AUTHORITY_HIERARCHY|CODE_CURRENCY|APPLICABILITY|PERMIT_REQUIREMENT|PERMIT_EXEMPTION|REVIEW_REQUIREMENT|PATHWAY|THRESHOLD|ENTITLEMENT|OTHER", "source_type": "code|ordinance|permit_page|checklist|application|interpretation|entitlement|other", "retrieval_note": "short", "rule": "specific proposition supported"}}],
"disciplines": [{{
"type": "string",
"scope_likelihood": "LIKELY|PROBABLE|POSSIBLE|UNLIKELY",
"scope_basis": "short SOW-based explanation",
"applicability": {{"rule": "short", "fact": {{"statement": "short", "source": "USER_PROVIDED|RETRIEVED_RECORD|AUTHORITATIVE_SOURCE|INFERRED|UNKNOWN"}}, "determination": "applies|does_not_apply|cannot_determine", "missing": "short", "relationship": "direct|conditional|not_established", "evidence": ["E1"]}},
"permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
"permit_finding": "short",
"permit_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
"pathway_basis_rule": "Only DIRECT_EVIDENCE, CONDITIONAL, or NOT_ESTABLISHED are valid basis values.",
"permit_evidence": ["E2"],
"pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
"pathway_finding": "short",
"pathway_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
"pathway_evidence": ["E3"],
"authority_evidence": ["E4"],
"missing": ["short"],
"reopen": ["short"],
"actionable_questions": ["specific question tied to SOW + AHJ + unresolved issue"],
"process_confidence": "ESTABLISHED|EXPECTED|NOT_ESTABLISHED",
"expected_process": "short source-grounded expected process",
"potential_issues": [{{"issue": "possible research lead", "why": "short SOW/research basis"}}]
}}]
}}
"""
                st.session_state.run_status = "Running — sending research request to Gemini"
                state_errors = validate_input_state_consistency({}, address, state)
                if state_errors:
                    st.session_state.debug_log = {
                        "status": "Input validation failed",
                        "validation_errors": state_errors,
                        "run_id": st.session_state.run_id,
                    }
                    st.session_state.error_msg = state_errors[0]
                    st.session_state.run_status = "Blocked — project metadata conflict"
                    result = {
                        "data": None,
                        "error": True,
                        "retry": False,
                        "msg": state_errors[0],
                        "debug": st.session_state.debug_log,
                    }
                else:
                    result = cached_gemini_call(prompt_hash, prompt)
                st.session_state.run_status = "Research response received — validating evidence" if not result.get("error") else "Gemini returned an error — displaying diagnostics"
                # Use a dedicated evidence-recovery pass before validation repair. This gives
                # Gemini another chance to find the actual permit rule without asking it to
                # regenerate a 20k+ token dossier. If recovery fails, the original dossier
                # remains authoritative and the normal deterministic/repair path continues.
                if isinstance(result.get("data"), dict):
                    # Boundary verification comes first. A postal city or conditional municipal identity
                    # must never be allowed to steer permit/process research to the wrong AHJ.
                    j = result["data"].get("jurisdiction") or {}
                    j_status = str(j.get("status") or "").upper()
                    j_city = str(j.get("city") or "").strip().lower()
                    if j_status != "VERIFIED" or j_city in {"", "not yet established", "unconfirmed"}:
                        j_prompt = build_jurisdiction_recovery_prompt(result["data"], address, state, project_date)
                        j_hash = hashlib.md5((prompt_hash + "|jurisdiction_recovery").encode()).hexdigest()
                        st.session_state.run_status = "Running — parcel jurisdiction verification"
                        j_result = cached_gemini_jurisdiction_recovery(j_hash, j_prompt)
                        result.setdefault("debug", {})["jurisdiction_recovery"] = j_result.get("debug", {})
                        if not j_result.get("error"):
                            result["data"] = merge_jurisdiction_recovery(result["data"], j_result.get("data") or {}, address)

                    # v27 deliberately does NOT run repeated permit/process recovery calls.
                    # The initial research pass is the research brief: jurisdiction, codes, scope-based
                    # discipline triage, authoritative findings, expected process, questions, and leads.
                    # This keeps a normal run affordable and avoids turning unresolved evidence into
                    # an expensive loop of near-duplicate searches.
                    result.setdefault("debug", {})["permit_recovery_targets"] = []
                    result.setdefault("debug", {})["permit_recovery_batches"] = []
                    result.setdefault("debug", {})["process_recovery_targets"] = []
                    result.setdefault("debug", {})["process_recovery_batches"] = []

                    # Re-run the deterministic contract after jurisdiction verification so the
                    # research brief cannot bypass the evidence firewall.
                    post_recovery_data, post_recovery_errors = run_deterministic_contract_pass(
                        result.get("data") or {}, address, project_date, state
                    )
                    result["data"] = post_recovery_data
                    if post_recovery_errors:
                        result["error"] = True
                        result["retry"] = False
                        result["msg"] = "Research completed, but the recovered dossier failed regulatory consistency validation."
                        result.setdefault("debug", {})["validation_errors"] = post_recovery_errors
                        result.setdefault("debug", {})["error_type"] = "Validation Failed"
                if result.get("error") and result.get("debug", {}).get("error_type") == "Validation Failed":
                    deterministic_data, deterministic_errors = run_deterministic_contract_pass(result.get("data") or {}, address, project_date, state)
                    if not deterministic_errors:
                        st.session_state.debug_log = {"first_attempt": result.get("debug", {}), "deterministic_repair": {"status": "success", "validation_errors": [], "note": "Deterministic contract repair passed; no paid Gemini validation-repair call was made."}}
                        st.session_state.report_data = deterministic_data
                        st.session_state.error_msg = None
                    else:
                        validation_errors = deterministic_errors
                        prior_json = json.dumps(deterministic_data, indent=2)
                        repair_prompt = f"""
    You are repairing an AHJ regulatory research dossier that was completed but failed deterministic consistency validation.
    Return ONLY the complete corrected JSON object. Do not explain the changes.
    PROJECT:
    State: {state} | Address: {address} | Date: {project_date}
    Type: {ptype} | Class: {bclass} | Entitlements: {existing_permit}
    SCOPE: {sow_text}
    VALIDATION ERRORS:
    {json.dumps(validation_errors, indent=2)}
    REPAIR CONTRACT:
    - Fix every validation error using ONLY the evidence and project facts already present in PREVIOUS JSON.
    - This is a STRUCTURAL/CONSISTENCY repair pass, not a research pass. Do NOT use Google Search and do NOT add new evidence records, URLs, citations, legal rules, permit rules, code rules, jurisdiction findings, pathway findings, or project facts.
    - Do not evade an error by deleting a discipline or evidence item merely to make validation pass.
    - JURISDICTION IS A HARD REQUIREMENT: use only existing jurisdiction evidence already present in PREVIOUS JSON. If it is insufficient, downgrade the jurisdiction rather than inventing site-specific evidence. Postal city/ZIP and generic agency coverage are not enough. If evidence establishes an unincorporated parcel, use "Unincorporated" for the actual city field.
    - CODE CURRENCY IS A HARD REQUIREMENT: use only existing CODE_CURRENCY evidence already present in PREVIOUS JSON. If it is insufficient, downgrade CURRENT to CONDITIONAL rather than inventing code-current evidence.
    - A code may be CURRENT only when valid CODE_CURRENCY evidence already establishes its edition and current/adopted/effective/mandatory status as of the project date.
    - If the cited evidence is for a newer edition than the code name, correct the code entry to the edition actually supported by the evidence; if current status remains unresolved, use CONDITIONAL. Do not preserve a stale code year merely to keep the original text.
    - Do not treat a prior edition being available online as evidence that it is current.
    - Preserve valid evidence and project facts.
    - If a permit requirement lacks discipline-matched PERMIT_REQUIREMENT evidence, downgrade the permit conclusion to CONDITIONAL/UNKNOWN. Do NOT create, relabel, or rewrite evidence to manufacture a permit requirement. A separate permit-recovery pass is responsible for researching missing permit evidence.
    - If a negative permit/exemption claim lacks discipline-matched PERMIT_EXEMPTION evidence, remove the definitive negative claim. Do NOT create or relabel evidence to manufacture an exemption. IMPORTANT: do not treat an epistemic statement such as "permit requirement is not established," "current evidence does not establish a permit requirement," or "cannot determine whether a permit is required" as a legal exemption. Those statements are allowed with NOT_ESTABLISHED / UNKNOWN / CONDITIONAL status and do not require PERMIT_EXEMPTION evidence.
    - If a conditional permit finding says a condition "requires" or "triggers" a permit, it still needs PERMIT_REQUIREMENT evidence; otherwise state that the permit consequence is not established.
    - REVIEW_REQUIREMENT establishes review/engineering/inspection obligations; it does NOT establish a permit requirement.
    - PATHWAY establishes process only; it does NOT establish a permit requirement.
    - APPLICABILITY and THRESHOLD evidence do NOT establish downstream permit/pathway consequences by themselves.
    - THRESHOLD evidence can explain a condition but cannot be relabeled as PERMIT_REQUIREMENT.
    - If an evidence item is labeled THRESHOLD, its rule itself must contain a threshold/limit proposition. If it does not, change the proposition to OTHER; do not invent a threshold.
    - Do not treat numbers appearing only in project facts, equipment specifications, titles, or unrelated source text as regulatory thresholds. The source rule must establish the threshold.
    - A statement that a permit is not required, not needed, or exempt is a legal exemption claim. It requires valid PERMIT_EXEMPTION evidence. Without that evidence, rewrite the statement as a non-establishment/unknown statement.
    - Do not convert a plausible workflow into a verified pathway without PATHWAY or REVIEW_REQUIREMENT evidence.
    - Any concrete pathway/process statement (portal, submit, file, processed, plan review, concurrently, separately) must have valid PATHWAY or REVIEW_REQUIREMENT evidence even when the pathway status is CONDITIONAL. Otherwise state that the pathway is not established.
    - Never invent missing project facts.
    - Never infer CUP conditions or amendment consequences without the governing entitlement or authoritative amendment rule.
    Never put a CUP amendment/no-amendment consequence into an applicability Source Rule unless the cited applicability evidence is proposition-valid ENTITLEMENT evidence.
    - Bottom Line may only summarize conclusions actually established in the discipline findings.
    - If the evidence is insufficient, say so explicitly rather than manufacturing certainty.
    PREVIOUS JSON:
    {prior_json}
    Use the same schema and proposition_type taxonomy as the original research contract.
    REPAIR RULE — DO NOT RELABEL EVIDENCE:
    If validation says a permit consequence lacks PERMIT_REQUIREMENT evidence, do not change the evidence label unless the source rule itself explicitly establishes a permit/approval requirement. If the source only establishes review, inspection, threshold, applicability, compliance, or pathway, preserve that proposition type and downgrade the permit conclusion to CONDITIONAL or UNKNOWN.
    - SOURCE-SPECIFICITY RULE: Generic agency homepages, service indexes, generic code indexes, and permit portals such as EPIC-LA are not proposition-specific support for PERMIT_REQUIREMENT, PERMIT_EXEMPTION, PATHWAY, or ENTITLEMENT.
    - EVIDENCE-CONTENT RULE: The evidence rule itself must state the claimed proposition; title/retrieval_note alone never proves it.
    - NO-NEW-EVIDENCE RULE: The repaired JSON must not contain any evidence ID that was absent from PREVIOUS JSON. Do not invent URLs, source titles, section numbers, quotations, or legal propositions.
    - NO-NEW-CONCLUSION-RULE: A definitive permit/pathway/jurisdiction/code conclusion may only remain if the corresponding valid evidence already exists in PREVIOUS JSON. Otherwise downgrade it.
    - EPISTEMIC RULE: “does not yet establish whether a permit is required,” “not yet established,” and “cannot determine whether a permit is required” are unresolved, not affirmative permit conclusions.
    """
                        repair_hash = hashlib.md5((
                            PROMPT_VERSION + "|validation_repair|" + prompt_hash + "|" +
                            json.dumps(validation_errors, sort_keys=True)
                        ).encode()).hexdigest()
                        st.session_state.run_status = "Running — deterministic validation repair"
                        repair_result = cached_gemini_repair(repair_hash, repair_prompt)
                        st.session_state.debug_log = {
                            "first_attempt": result.get("debug", {}),
                            "validation_repair": repair_result.get("debug", {}),
                        }
                        if repair_result.get("error"):
                            st.session_state.error_msg = repair_result.get("msg", "Validation repair failed.")
                            st.session_state.report_data = None
                        else:
                            repaired_data = constrain_validation_repair_output(
                                repair_result.get("data") or {}, deterministic_data
                            )
                            repaired_data, repaired_errors = run_deterministic_contract_pass(
                                repaired_data, address, project_date, state
                            )
                            if repaired_errors:
                                repair_debug = repair_result.get("debug", {})
                                repair_debug["validation_errors"] = repaired_errors
                                repair_debug["error_type"] = "Validation Failed After Repair"
                                st.session_state.debug_log["validation_repair"] = repair_debug
                                st.session_state.error_msg = "Dossier still failed regulatory consistency validation after repair. Gemini repair could not satisfy the deterministic evidence contract."
                                st.session_state.report_data = None
                            else:
                                repair_debug = repair_result.get("debug", {})
                                repair_debug["validation_errors"] = []
                                repair_debug["deterministic_post_repair"] = "passed"
                                st.session_state.debug_log["validation_repair"] = repair_debug
                                st.session_state.report_data = repaired_data
                                st.session_state.error_msg = None
                                st.session_state.run_status = "Complete — dossier passed validation"
                elif result.get("retry"):
                    retry_prompt = f"""
You are completing a regulatory research dossier that previously hit the generation limit.
Return ONLY the required JSON object.
IMPORTANT:
- Do not redo unnecessary research. Preserve authoritative findings already established.
- Do not omit a discipline merely to save tokens. Do not invent missing facts.
- Do not add explanatory prose outside JSON. Keep wording extremely concise.
- Evidence rules must remain proposition-specific. Applicability, permit requirement, and pathway must remain separate.
- Preserve all material evidence IDs. If something cannot be established, use UNKNOWN or CONDITIONAL.
- Never use NOT_APPLICABLE as a substitute for UNKNOWN.
COMPRESSION RULES:
- bottom_line: maximum 3 short sentences
- rule: maximum 15 words
- fact.statement: maximum 12 words
- permit_finding: maximum 18 words
- pathway_finding: maximum 18 words
- retrieval_note: maximum 12 words
- reason: maximum 20 words
- missing/reopen/critical_missing items: short phrases
- Do not repeat evidence text across multiple fields.
PROJECT:
State: {state} | Address: {address} | Date: {project_date}
Type: {ptype} | Class: {bclass} | Entitlements: {existing_permit}
SCOPE: {sow_text}
CODE CURRENCY FIREWALL:
Re-verify every CURRENT code against authoritative adoption/current-code evidence as of the project date.
AUTHORITY-HIERARCHY FIREWALL: Re-verify whether state rules control this jurisdiction for each material permit conclusion. If permit evidence is state-level, use AUTHORITY_HIERARCHY evidence when research identifies a material state/local control issue; otherwise a separate authority_evidence link is not required. If local authority modifies or controls, use the local rule.
Do not rely on model memory. A prior edition being available online does not make it CURRENT.
If the current edition cannot be established, use CONDITIONAL rather than asserting CURRENT.
Code name/year must agree with its CODE_CURRENCY evidence.
JSON SCHEMA:
{{
"bottom_line": "3 concise sentences maximum",
"bottom_line_evidence": ["E1"],
"research_completeness": {{"status": "SUFFICIENT|PARTIAL|INSUFFICIENT", "reason": "short", "critical_missing": ["short item"]}},
"jurisdiction": {{"status": "VERIFIED|CONDITIONAL", "county": "string", "city": "string", "ahj": "string", "evidence": ["E1"]}},
"codes": [{{"name": "string", "status": "CURRENT|CONDITIONAL", "evidence": ["E2"]}}],
"evidence": [{{"id": "E1", "title": "string", "url": "string", "authority": "state|county|city|federal|tribal|other", "discipline": "string", "proposition_type": "JURISDICTION|AUTHORITY_HIERARCHY|CODE_CURRENCY|APPLICABILITY|PERMIT_REQUIREMENT|PERMIT_EXEMPTION|REVIEW_REQUIREMENT|PATHWAY|THRESHOLD|ENTITLEMENT|OTHER", "source_type": "code|ordinance|permit_page|checklist|application|interpretation|entitlement|other", "retrieval_note": "short", "rule": "specific proposition supported"}}],
"disciplines": [{{
"type": "string",
"applicability": {{"rule": "short", "fact": {{"statement": "short", "source": "USER_PROVIDED|RETRIEVED_RECORD|AUTHORITATIVE_SOURCE|INFERRED|UNKNOWN"}}, "determination": "applies|does_not_apply|cannot_determine", "missing": "short", "relationship": "direct|conditional|not_established", "evidence": ["E1"]}},
"permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
"permit_finding": "short",
"permit_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
"pathway_basis_rule": "Only DIRECT_EVIDENCE, CONDITIONAL, or NOT_ESTABLISHED are valid basis values.",
"permit_evidence": ["E2"],
"pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
"pathway_finding": "short",
"pathway_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
"pathway_evidence": ["E3"],
"authority_evidence": ["E4"],
"missing": ["short"],
"reopen": ["short"],
"actionable_questions": ["specific question tied to SOW + AHJ + unresolved issue"],
"potential_issues": [{{"issue": "possible research lead", "why": "short SOW/research basis"}}]
}}]
}}
"""
                    retry_result = cached_gemini_retry(prompt_hash + "_retry", retry_prompt)
                    if retry_result.get("error"):
                        st.session_state.debug_log = {
                            "first_attempt": result.get("debug", {}),
                            "retry_attempt": retry_result.get("debug", {}),
                        }
                        st.session_state.error_msg = retry_result.get("msg", "Retry failed.")
                        st.session_state.run_status = "Failed — retry did not complete"
                    else:
                        st.session_state.debug_log = {
                            "first_attempt": result.get("debug", {}),
                            "retry_attempt": retry_result.get("debug", {}),
                        }
                        st.session_state.report_data = retry_result["data"]
                        st.session_state.run_status = "Complete — dossier passed validation"
                else:
                    st.session_state.debug_log = result.get("debug", {})
                    if result["error"]:
                        st.session_state.error_msg = result["msg"]
                        st.session_state.run_status = "Failed — Gemini/API error"
                    else:
                        st.session_state.report_data = result["data"]
                        st.session_state.run_status = "Complete — dossier passed validation"

st.info(f"Research status: {st.session_state.run_status}")
if st.session_state.run_id:
    st.caption(f"Run ID: {st.session_state.run_id} · started {st.session_state.run_started_utc}")

if st.session_state.error_msg:
    st.error(f"❌ {st.session_state.error_msg}")

validation_errors = st.session_state.debug_log.get("validation_errors", []) if isinstance(st.session_state.debug_log, dict) else []
repair_errors = st.session_state.debug_log.get("validation_repair", {}).get("validation_errors", []) if isinstance(st.session_state.debug_log, dict) else []
visible_errors = repair_errors or validation_errors

if visible_errors:
    st.markdown("**Why the dossier was blocked:**")
    with st.expander("⚠️ View Regulatory Consistency Errors", expanded=True):
        for err in visible_errors:
            st.warning(err)

with st.expander("💰 Gemini API Usage", expanded=bool(st.session_state.error_msg)):
    debug = st.session_state.debug_log if isinstance(st.session_state.debug_log, dict) else {}
    usage_rows = []
    for label, block in (("Initial research", debug), ("Jurisdiction recovery", debug.get("jurisdiction_recovery", {})), ("Generation retry", debug.get("retry_attempt", {})), ("Permit recovery", debug.get("permit_recovery", {})), ("Process recovery", debug.get("process_recovery", {})), ("Validation repair", debug.get("validation_repair", {}))):
        usage = block.get("api_usage") if isinstance(block, dict) else None
        if isinstance(usage, dict):
            usage_rows.append({
                "call": label,
                "timestamp (UTC)": usage.get("timestamp_utc"),
                "model": usage.get("model"),
                "input": usage.get("input_tokens"),
                "output": usage.get("output_tokens"),
                "thoughts": usage.get("thoughts_tokens"),
                "total": usage.get("total_tokens"),
                "cached": usage.get("cached_tokens"),
                "seconds": usage.get("elapsed_seconds"),
            })
    # Legacy recovery usage rows are retained for cached older runs. v27 does not execute
    # repeated permit/process recovery batches.
    for phase_key, phase_label in (("permit_recovery_batches", "Permit recovery"), ("process_recovery_batches", "Process recovery")):
        for batch in debug.get(phase_key, []) if isinstance(debug.get(phase_key), list) else []:
            usage = (batch.get("debug") or {}).get("api_usage") if isinstance(batch, dict) else None
            if isinstance(usage, dict):
                usage_rows.append({
                    "call": f"{phase_label} #{batch.get('batch', '?')}",
                    "timestamp (UTC)": usage.get("timestamp_utc"),
                    "model": usage.get("model"),
                    "input": usage.get("input_tokens"),
                    "output": usage.get("output_tokens"),
                    "thoughts": usage.get("thoughts_tokens"),
                    "total": usage.get("total_tokens"),
                    "cached": usage.get("cached_tokens"),
                    "seconds": usage.get("elapsed_seconds"),
                })
    if usage_rows:
        st.dataframe(usage_rows, use_container_width=True, hide_index=True)
        st.caption("These numbers come directly from Gemini usage_metadata for each actual API call. Successful research results are cached for 1 hour; failed API calls are never cached.")
    else:
        st.caption("No Gemini API call has been recorded for the current run.")
    if os.path.exists(USAGE_LOG_PATH):
        try:
            with open(USAGE_LOG_PATH, "rb") as _usage_file:
                st.download_button(
                    "Download full usage log",
                    data=_usage_file.read(),
                    file_name="gemini_usage.log",
                    mime="application/jsonl",
                    use_container_width=True,
                )
        except Exception:
            pass

with st.expander("🐛 API Debug Log", expanded=bool(st.session_state.error_msg)):
    st.json(st.session_state.debug_log)

# ============================================================
# USER-FACING RESULTS / EXPORT
# ============================================================
def _pretty_fact_source(value):
    labels = {
        "USER_PROVIDED": "Project scope",
        "RETRIEVED_RECORD": "Project record",
        "AUTHORITATIVE_SOURCE": "Authoritative source",
        "INFERRED": "AI inference",
        "UNKNOWN": "Not established",
    }
    return labels.get(str(value or "").upper(), str(value or "Not established").replace("_", " ").title())

def _pretty_status(value):
    labels = {
        "VERIFIED_REQUIRED": "Permit required",
        "NOT_APPLICABLE": "Not applicable",
        "NOT_CURRENTLY_TRIGGERED": "Not currently triggered",
        "CONDITIONAL": "Conditional",
        "UNKNOWN": "Not yet established",
        "INFERRED": "Inference only",
        "USER_PROVIDED": "From project scope",
    }
    return labels.get(str(value or "").upper(), str(value or "Not established").replace("_", " ").title())

def _pretty_scope_likelihood(value):
    return {
        "LIKELY": "Likely",
        "PROBABLE": "Probable",
        "POSSIBLE": "Possible",
        "UNLIKELY": "Unlikely",
    }.get(str(value or "").upper(), "Possible")

def _pretty_process_status(value):
    return {
        "PATHWAY_ESTABLISHED": "Pathway established",
        "REVIEW_REQUIRED": "Review requirement established",
        "NOT_ESTABLISHED": "Not yet established",
    }.get(str(value or "").upper(), "Not yet established")

def _pretty_pathway_status(value):
    labels = {
        "VERIFIED_REQUIRED": "Established",
        "NOT_APPLICABLE": "Not applicable",
        "NOT_CURRENTLY_TRIGGERED": "Not currently triggered",
        "CONDITIONAL": "Conditional",
        "UNKNOWN": "Not yet established",
        "INFERRED": "Inference only",
        "USER_PROVIDED": "From project scope",
    }
    return labels.get(str(value or "").upper(), str(value or "Not established").replace("_", " ").title())

def _pretty_applicability(value):
    labels = {
        "applies": "Applies",
        "does_not_apply": "Does not apply",
        "cannot_determine": "Needs more information",
    }
    return labels.get(str(value or "").lower(), str(value or "Not established").replace("_", " ").title())

def _status_icon(value):
    return {
        "VERIFIED_REQUIRED": "🟢",
        "CONDITIONAL": "🟠",
        "UNKNOWN": "🟡",
        "NOT_APPLICABLE": "⚪",
        "NOT_CURRENTLY_TRIGGERED": "⚪",
        "INFERRED": "🟣",
        "USER_PROVIDED": "🔵",
    }.get(str(value or "").upper(), "🟡")

def sanitize_generic_actionable_questions(data):
    if not isinstance(data, dict):
        return data
    generic_patterns = [
        r"what code or permit provision determines whether the work needs a permit",
        r"what review or submission route does this ahj use for this specific scope",
        r"does this exact .* scope require (a )?separate permit",
        r"does this .* scope require (a )?permit",
        r"are there any permit exemptions that apply",
        r"if (a )?permit (or approval)? is required, where is the application submitted",
        r"confirm the permit consequence",
        r"confirm the applicable submission process",
    ]
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        raw = item.get("actionable_questions") or []
        if isinstance(raw, str):
            raw = [raw]
        kept = []
        for q in raw:
            q = str(q).strip()
            if not q:
                continue
            qn = _norm_text(q)
            if any(re.search(pattern, qn) for pattern in generic_patterns):
                continue
            if q not in kept:
                kept.append(q)
        item["actionable_questions"] = kept[:5]
    return data

def _soften_inference_lead(text):
    text = re.sub(r"\bmay(?:\s+may)+\b", "may", str(text or "").strip(), flags=re.I)
    if not text:
        return text
    pattern = re.compile(
        r"(?<!may )(?<!may be )\b(?:will\s+be\s+required|will\s+require|will\s+trigger|must|is\s+required|are\s+required|requires|require|triggers|trigger|required)\b",
        re.I,
    )
    def repl(match):
        phrase = match.group(0).lower()
        if phrase.startswith("will be required") or phrase == "required":
            return "may be required"
        if phrase.startswith("will require") or phrase in {"requires", "require"}:
            return "may require"
        if phrase.startswith("will trigger") or phrase in {"triggers", "trigger"}:
            return "may trigger"
        if phrase == "must":
            return "may need to"
        if phrase in {"is required", "are required"}:
            return "may be required"
        return match.group(0)
    return pattern.sub(repl, text)

def sanitize_generated_prose(data):
    """Fix a few obvious Gemini wording artifacts without changing regulatory meaning."""
    replacements = {
        "could may require": "may require",
        "could might require": "might require",
        "may might require": "may require",
        "could may warrant": "may warrant",
        "could might warrant": "might warrant",
        "may might warrant": "may warrant",
        "a electrical": "an electrical",
        "a energy": "an energy",
        "a architectural": "an architectural",
    }

    def clean(value):
        if isinstance(value, str):
            out = value
            for old, new in replacements.items():
                out = re.sub(re.escape(old), new, out, flags=re.IGNORECASE)
            return out
        if isinstance(value, list):
            return [clean(v) for v in value]
        if isinstance(value, dict):
            return {k: clean(v) for k, v in value.items()}
        return value

    return clean(data)


def sanitize_potential_issues(data):
    if not isinstance(data, dict):
        return data
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        raw = item.get("potential_issues", [])
        if isinstance(raw, str):
            raw = [raw]
        cleaned = []
        for entry in raw or []:
            if isinstance(entry, dict):
                issue = _soften_inference_lead(entry.get("issue", ""))
                why = str(entry.get("why", "")).strip()
                if not issue:
                    continue

                # The lead itself is intentionally non-regulatory, but Gemini can
                # put an unsupported legal conclusion into the explanatory "why".
                # That turns a harmless research lead into an implied rule claim.
                # Keep the lead useful while removing definitive legal language.
                why_markers = re.compile(
                    r"\b(?:requires?|required|must|shall|prohibits?|prohibited|exempts?|exempt|"
                    r"triggers?|triggered|is mandated|are mandated|is required|are required)\b",
                    re.I,
                )
                if why_markers.search(why):
                    why = "This is a research lead based on the stated scope; confirm the applicable AHJ rule or project-specific condition."
                cleaned.append({"issue": issue, "why": why})
            else:
                text = _soften_inference_lead(entry)
                if text:
                    cleaned.append({"issue": text, "why": ""})
        seen = set()
        final = []
        for entry in cleaned:
            key = _norm_text(entry.get("issue"))
            if key and key not in seen:
                seen.add(key)
                final.append(entry)
        item["potential_issues"] = final[:5]
    return data

def _decision_reference(item, ev_dict):
    app = item.get("applicability") or {}
    permit = str(item.get("permit", "UNKNOWN")).upper()
    pathway = str(item.get("pathway", "UNKNOWN")).upper()
    def ids(value):
        if isinstance(value, str):
            value = [value]
        return [str(v) for v in (value or []) if str(v) in ev_dict]
    app_ids = ids(app.get("evidence", []))
    permit_ids = ids(item.get("permit_evidence", []))
    pathway_ids = ids(item.get("pathway_evidence", []))
    fact = app.get("fact", {})
    fact_statement = (fact.get("statement", "") if isinstance(fact, dict) else str(fact)).strip()
    app_rule = str(app.get("rule", "")).strip() if app_ids else ""
    # A permit decision must display the rule from the SAME validated permit
    # evidence that caused the deterministic decision. Never display the
    # broader applicability rule here; doing so makes a valid permit decision
    # look as though it was inferred from mere code applicability.
    permit_rule = ""
    if permit_ids:
        permit_rules = []
        for eid in permit_ids:
            ev = ev_dict.get(eid) or {}
            rule = str(ev.get("rule") or "").strip()
            if rule and ev.get("proposition_type") in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION"}:
                permit_rules.append(rule)
        permit_rule = permit_rules[0] if permit_rules else ""

    if permit == "VERIFIED_REQUIRED":
        refs = permit_ids
        trail_type = "Permit decision reference"
        conclusion = str(item.get("permit_finding", "")).strip()
        app_rule = permit_rule
    elif permit == "NOT_APPLICABLE":
        refs = permit_ids
        trail_type = "Permit decision reference"
        conclusion = str(item.get("permit_finding", "")).strip()
        app_rule = permit_rule
    elif pathway == "VERIFIED_REQUIRED":
        refs = pathway_ids
        trail_type = "Review / pathway decision reference"
        conclusion = str(item.get("pathway_finding", "")).strip()
    elif pathway == "NOT_APPLICABLE":
        refs = pathway_ids
        trail_type = "Review / pathway decision reference"
        conclusion = str(item.get("pathway_finding", "")).strip()
    else:
        refs = app_ids
        trail_type = "Applicability reference"
        conclusion = ""
    refs = list(dict.fromkeys(refs))[:3]
    return {
        "evidence_ids": refs,
        "fact": fact_statement,
        "rule": app_rule,
        "conclusion": conclusion,
        "has_authoritative_reference": bool(refs),
        "type": trail_type,
        "applicability_evidence_ids": app_ids[:3],
        "permit_evidence_ids": permit_ids[:3],
        "pathway_evidence_ids": pathway_ids[:3],
    }

def _plain_language_summary(item):
    permit = str(item.get("permit", "UNKNOWN")).upper()
    pathway = str(item.get("pathway", "UNKNOWN")).upper()
    app = item.get("applicability") or {}
    determination = str(app.get("determination", "")).lower()
    if permit == "VERIFIED_REQUIRED":
        permit_line = "A permit is required based on the cited permit-specific evidence."
    elif permit == "NOT_APPLICABLE":
        permit_line = "No permit is required based on the cited exemption evidence."
    elif permit == "NOT_CURRENTLY_TRIGGERED":
        permit_line = "No current permit trigger was established; this should be revisited if the project facts change."
    elif permit == "CONDITIONAL":
        permit_line = "The permit outcome depends on a specific unresolved project fact."
    else:
        permit_line = "A permit requirement has not yet been established by the available evidence."
    if pathway == "VERIFIED_REQUIRED":
        pathway_line = "The review/submission pathway is established."
    elif pathway == "NOT_APPLICABLE":
        pathway_line = "No separate review/submission pathway applies based on the cited evidence."
    elif pathway == "CONDITIONAL":
        pathway_line = "The review/submission pathway depends on an unresolved fact."
    else:
        pathway_line = "The review/submission pathway has not yet been established."
    if determination == "applies":
        apply_line = "This discipline applies to the stated scope."
    elif determination == "does_not_apply":
        apply_line = "This discipline does not apply to the stated scope."
    else:
        apply_line = "Whether this discipline applies still needs to be confirmed."
    return apply_line, permit_line, pathway_line

def _discipline_questions(item):
    supplied = item.get("actionable_questions", [])
    if isinstance(supplied, str):
        supplied = [supplied]
    if isinstance(supplied, list):
        cleaned = []
        for value in supplied:
            text = str(value).strip()
            if text and text not in cleaned:
                cleaned.append(text)
        if cleaned:
            return cleaned[:5]
    discipline = str(item.get("type", "This discipline"))
    permit = str(item.get("permit", "UNKNOWN")).upper()
    pathway = str(item.get("pathway", "UNKNOWN")).upper()
    app = item.get("applicability") or {}
    determination = str(app.get("determination", "")).lower()
    questions = []
    fact = app.get("fact", {})
    fact_text = fact.get("statement", "") if isinstance(fact, dict) else str(fact)
    rule_text = str(app.get("rule", ""))
    missing = item.get("missing", [])
    if isinstance(missing, str):
        missing = [missing]
    missing_text = [str(v).strip() for v in missing if str(v).strip()]
    if determination in {"cannot_determine", "unknown"} and missing_text:
        questions.append(
            f"For the stated {discipline.lower()} scope ({fact_text}), what specific fact or document does the AHJ need to determine whether the {rule_text or 'applicable requirement'} applies?"
        )
    if permit in {"UNKNOWN", "CONDITIONAL"} and missing_text:
        questions.append(
            f"For this {discipline.lower()} scope, can the AHJ confirm whether {missing_text[0].rstrip('.')} changes the permit determination?"
        )
    for value in missing_text[1:]:
        if len(questions) >= 5:
            break
        questions.append(f"Can the project team confirm {value} for this scope?")
    return questions[:5]

def _add_doc_title(doc, title, subtitle=None):
    p = doc.add_paragraph()
    p.style = doc.styles["Title"]
    run = p.add_run(title)
    run.bold = True
    if subtitle:
        p2 = doc.add_paragraph(subtitle)
        p2.style = doc.styles["Subtitle"]

def _shade_cell(cell, fill):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)

def _set_doc_margins(section):
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

if st.session_state.report_data:
    data = sanitize_generated_prose(st.session_state.report_data)
    data = sanitize_scope_relevance(data)
    data = sanitize_expected_process(data)
    st.session_state.report_data = data
    data = sanitize_generic_actionable_questions(st.session_state.report_data)
    data = sanitize_potential_issues(st.session_state.report_data)
    st.session_state.report_data = data
    st.divider()
    st.header("4. Research Dossier")
    if "validation_errors" in st.session_state.debug_log:
        st.error("⚠️ Regulatory consistency validation failed.")
        with st.expander("View validation errors", expanded=True):
            for err in st.session_state.debug_log["validation_errors"]:
                st.markdown(f"- {err}")
    repair_debug = st.session_state.debug_log.get("validation_repair", {}) if isinstance(st.session_state.debug_log, dict) else {}
    if repair_debug.get("validation_errors"):
        st.error("⚠️ One-pass self-correction also failed validation.")
        with st.expander("View post-repair validation errors", expanded=True):
            for err in repair_debug.get("validation_errors", []):
                st.markdown(f"- {err}")
    st.info(f"**Bottom Line:** {data.get('bottom_line', 'N/A')}")
    bl_ev = data.get("bottom_line_evidence", [])
    if bl_ev:
        st.caption(f"Evidence IDs: {', '.join(bl_ev)}")
    completeness = data.get("research_completeness") or {}
    if completeness:
        completeness_status = completeness.get("status", "UNKNOWN")
        if completeness_status == "SUFFICIENT":
            st.success("Research completeness: SUFFICIENT")
        elif completeness_status == "PARTIAL":
            st.warning(f"Research completeness: PARTIAL — {completeness.get('reason', 'Material items remain unresolved.')}")
        elif completeness_status == "INSUFFICIENT":
            st.error(f"Research completeness: INSUFFICIENT — {completeness.get('reason', 'Authoritative evidence is insufficient.')}")
        critical_missing = completeness.get("critical_missing", [])
        if critical_missing:
            st.markdown("**Information that would help finish the research:**")
            for item in critical_missing:
                st.markdown(f"- {item}")
    st.subheader("📍 Jurisdiction")
    jur = data.get("jurisdiction") or {}
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**County:** {jur.get('county', 'Unknown')}")
        st.write(f"**City:** {jur.get('city', 'Unknown')}")
        st.write(f"**AHJ:** {jur.get('ahj', 'Unknown')}")
    with col2:
        ev_dict = {ev["id"]: ev for ev in data.get("evidence", [])}
        for eid in jur.get("evidence", []):
            if eid in ev_dict:
                ev = ev_dict[eid]
                st.write(f"**Source:** [{ev['title']}]({ev['url']})")
    st.subheader("📚 Applicable Codes")
    for code in (data.get("codes") or []):
        code_name = code.get("name", "Unknown")
        code_status = code.get("status", "N/A")
        st.markdown(f"- **{code_name}** — {_pretty_status(code_status)}")
    st.subheader("📋 What We Found")
    st.caption("The technical evidence remains available below, but the main view uses plain-language labels.")
    for item in data.get("disciplines", []):
        permit = str(item.get("permit", "UNKNOWN")).upper()
        pathway = str(item.get("pathway", "UNKNOWN")).upper()
        app = item.get("applicability") or {}
        determination = app.get("determination", "")
        icon = _status_icon(permit)
        title = item.get("type", "Unknown")
        permit_label = _pretty_status(permit)
        pathway_label = _pretty_pathway_status(pathway)
        apply_line, permit_line, pathway_line = _plain_language_summary(item)
        with st.expander(f"{icon} {title} — {_pretty_scope_likelihood(item.get('scope_likelihood'))} scope signal", expanded=False):
            if item.get("scope_basis"):
                st.caption(f"Why it is on the research map: {item.get('scope_basis')}")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Scope signal", _pretty_scope_likelihood(item.get("scope_likelihood")))
            with c2:
                st.metric("Permit", permit_label)
            with c3:
                st.metric("Review / pathway", pathway_label)
            st.markdown("### In plain English")
            st.markdown(f"**Permit:** {permit_line}")
            st.markdown(f"**Review:** {pathway_line}")
            process_finding = str(item.get("process_finding") or "").strip()
            process_ids = item.get("process_evidence") or []
            if process_finding:
                st.markdown(f"**Process evidence:** {process_finding}")
            expected_process = str(item.get("expected_process") or "").strip()
            process_confidence = str(item.get("process_confidence") or "NOT_ESTABLISHED").upper()
            if expected_process:
                label = "Expected AHJ process" if process_confidence == "EXPECTED" else "AHJ process"
                st.markdown(f"**{label}:** {expected_process}")
            if process_ids:
                process_details = [
                    ("Who reviews", item.get("process_reviewer")),
                    ("How it moves", item.get("process_route")),
                    ("Relationship", item.get("process_relationship")),
                    ("What changes it", item.get("process_trigger")),
                ]
                for label, value in process_details:
                    value = str(value or "").strip()
                    if value:
                        st.markdown(f"**{label}:** {value}")
                process_sources = []
                for eid in process_ids:
                    ev = ev_dict.get(str(eid))
                    if ev and ev.get("url"):
                        process_sources.append(f"[{ev.get('title', eid)}]({ev.get('url')})")
                if process_sources:
                    st.caption("Process source(s): " + "; ".join(process_sources))
            fact = app.get("fact", {})
            fact_statement = fact.get("statement", "N/A") if isinstance(fact, dict) else str(fact)
            fact_source = fact.get("source", "UNKNOWN") if isinstance(fact, dict) else "UNKNOWN"
            questions = _discipline_questions(item)
            potential_issues = item.get("potential_issues") or []
            missing = item.get("missing", [])
            if isinstance(missing, str):
                missing = [missing]
            reopen = item.get("reopen", [])
            if isinstance(reopen, str):
                reopen = [reopen]
            if questions:
                st.markdown("### What should we resolve next?")
                for q in questions:
                    st.markdown(f"- {q}")
            if potential_issues:
                st.markdown("### 💡 Worth checking — AI-generated possible lead")
                for lead in potential_issues:
                    if isinstance(lead, dict):
                        issue = str(lead.get("issue", "")).strip()
                        why = str(lead.get("why", "")).strip()
                    else:
                        issue, why = str(lead).strip(), ""
                    if issue:
                        st.markdown(f"- {issue}")
                    if why:
                        st.caption(f"Why it came up: {why}")
            with st.expander("Research trail & source details", expanded=False):
                decision_ref = _decision_reference(item, ev_dict)
                if decision_ref["fact"]:
                    st.markdown(f"**Project fact used:** {decision_ref['fact']}")
                if decision_ref["rule"]:
                    st.markdown(f"**Rule/source finding:** {decision_ref['rule']}")
                if decision_ref["conclusion"]:
                    st.markdown(f"**Conclusion:** {decision_ref['conclusion']}")
                if decision_ref["has_authoritative_reference"]:
                    st.caption("Supporting source(s):")
                    for eid in decision_ref["evidence_ids"]:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})** — {ev.get('proposition_type', 'OTHER')}")
                elif decision_ref["applicability_evidence_ids"]:
                    st.info("The cited source supports applicability only. It is not being treated as proof of a permit requirement or pathway.")
                    for eid in decision_ref["applicability_evidence_ids"][:3]:
                        if eid in ev_dict:
                            ev = ev_dict[eid]
                            st.markdown(f"- **[{ev['title']}]({ev['url']})** — applicability evidence")
                else:
                    st.caption("No proposition-specific authoritative evidence was established for this decision.")
                st.markdown(f"**Project information:** {fact_statement} [{_pretty_fact_source(fact_source)}]")
                permit_finding = str(item.get("permit_finding", "N/A"))
                pathway_finding = str(item.get("pathway_finding", "N/A"))
                st.markdown(f"**Permit finding:** {permit_finding}")
                st.markdown(f"**Review / pathway finding:** {pathway_finding}")
                app_ev = app.get("evidence", [])
                permit_ev = item.get("permit_evidence", [])
                pathway_ev = item.get("pathway_evidence", [])
                if app_ev or permit_ev or pathway_ev:
                    st.markdown("**Evidence by proposition**")
                    if app_ev:
                        st.write("Applicability")
                        for eid in app_ev:
                            if eid in ev_dict:
                                ev = ev_dict[eid]
                                st.caption(f"{ev.get('title', eid)} — {ev.get('rule', 'N/A')}")
                    if permit_ev:
                        st.write("Permit")
                        for eid in permit_ev:
                            if eid in ev_dict:
                                ev = ev_dict[eid]
                                st.caption(f"{ev.get('title', eid)} — {ev.get('rule', 'N/A')}")
                    if pathway_ev:
                        st.write("Review / pathway")
                        for eid in pathway_ev:
                            if eid in ev_dict:
                                ev = ev_dict[eid]
                                st.caption(f"{ev.get('title', eid)} — {ev.get('rule', 'N/A')}")
                                if ev.get("proposition_type") in {"PATHWAY", "REVIEW_REQUIREMENT"}:
                                    st.caption(f"Process evidence: {ev.get('proposition_type')}")
                process_ids = item.get("process_evidence") or []
                if process_ids:
                    st.markdown("**Process evidence**")
                    for eid in process_ids:
                        if eid in ev_dict:
                            ev = ev_dict[eid]
                            st.caption(f"{ev.get('title', eid)} — {ev.get('proposition_type', 'OTHER')}: {ev.get('rule', 'N/A')}")
                if missing:
                    st.markdown("**Missing information**")
                    for value in missing:
                        st.markdown(f"- {value}")
                if reopen:
                    st.markdown("**Reopen if**")
                    for value in reopen:
                        st.markdown(f"- {value}")
    st.header("5. Export")
    col1, col2 = st.columns(2)
    with col1:
        doc = Document()
        for section in doc.sections:
            _set_doc_margins(section)
        doc.styles["Normal"].font.name = "Aptos"
        doc.styles["Normal"].font.size = Pt(10)
        _add_doc_title(doc, "AHJ Research Dossier", f"{address} · {state} · {project_date}")
        doc.add_paragraph(f"Generated {datetime.now().strftime('%B %d, %Y')}")
        doc.add_heading("Bottom Line", level=1)
        p = doc.add_paragraph(data.get("bottom_line", ""))
        p.style = doc.styles["Normal"]
        doc.add_heading("Jurisdiction", level=1)
        table = doc.add_table(rows=4, cols=2)
        table.style = "Light Shading Accent 1"
        rows = [
            ("Status", jur.get("status", "N/A")),
            ("County", jur.get("county", "N/A")),
            ("City", jur.get("city", "N/A")),
            ("AHJ", jur.get("ahj", "N/A")),
        ]
        for row, (label, value) in zip(table.rows, rows):
            row.cells[0].text = label
            row.cells[1].text = str(value)
        doc.add_heading("Applicable Codes", level=1)
        for code in (data.get("codes") or []):
            doc.add_paragraph(f"{code.get('name')} — {_pretty_status(code.get('status'))}", style="List Bullet")
        doc.add_heading("Permit & Review Summary", level=1)
        matrix = doc.add_table(rows=1, cols=4)
        matrix.style = "Light Shading Accent 1"
        hdr = matrix.rows[0].cells
        for cell, text in zip(hdr, ["Discipline", "Scope signal", "Permit", "Review / pathway"]):
            cell.text = text
            _shade_cell(cell, "D9EAF7")
        for item in data.get("disciplines", []):
            row = matrix.add_row().cells
            row[0].text = str(item.get("type", "Unknown"))
            row[1].text = _pretty_scope_likelihood(item.get("scope_likelihood"))
            row[2].text = _pretty_status(item.get("permit"))
            row[3].text = _pretty_pathway_status(item.get("pathway"))
        doc.add_page_break()
        doc.add_heading("Discipline Findings", level=1)
        for item in data.get("disciplines", []):
            doc.add_heading(str(item.get("type", "Unknown")), level=2)
            p = doc.add_paragraph()
            p.add_run("Scope signal: ").bold = True
            p.add_run(_pretty_scope_likelihood(item.get("scope_likelihood")))
            if item.get("scope_basis"):
                p = doc.add_paragraph(str(item.get("scope_basis")))
            app = item.get("applicability") or {}
            fact = app.get("fact", {})
            fact_statement = fact.get("statement", "") if isinstance(fact, dict) else str(fact)
            fact_source = fact.get("source", "UNKNOWN") if isinstance(fact, dict) else "UNKNOWN"
            decision_ref = _decision_reference(item, ev_dict)
            doc.add_heading("At a glance", level=3)
            apply_line, permit_line, pathway_line = _plain_language_summary(item)
            glance = doc.add_table(rows=1, cols=2)
            glance.style = "Table Grid"
            for cell, text in zip(glance.rows[0].cells, ["Permit", "Review / pathway"]):
                cell.text = text
            row = glance.add_row().cells
            row[0].text = _pretty_status(item.get("permit"))
            row[1].text = _pretty_pathway_status(item.get("pathway"))
            p = doc.add_paragraph()
            p.add_run("Permit: ").bold = True
            p.add_run(permit_line)
            p = doc.add_paragraph()
            p.add_run("Review: ").bold = True
            p.add_run(pathway_line)
            process_finding = str(item.get("process_finding") or "").strip()
            if process_finding:
                p = doc.add_paragraph()
                p.add_run("Process evidence: ").bold = True
                p.add_run(process_finding)
            expected_process = str(item.get("expected_process") or "").strip()
            if expected_process:
                p = doc.add_paragraph()
                p.add_run("Expected AHJ process: ").bold = True
                p.add_run(expected_process)
            process_details = [
                ("Who reviews", item.get("process_reviewer")),
                ("How it moves", item.get("process_route")),
                ("Relationship", item.get("process_relationship")),
                ("What changes it", item.get("process_trigger")),
            ]
            for label, value in process_details:
                value = str(value or "").strip()
                if value:
                    p = doc.add_paragraph()
                    p.add_run(f"{label}: ").bold = True
                    p.add_run(value)
            process_ids = item.get("process_evidence") or []
            if process_ids:
                doc.add_paragraph("Process source(s):")
                for eid in process_ids:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        p = doc.add_paragraph(style="List Bullet")
                        p.add_run(ev.get("title", eid))
                        if ev.get("url"):
                            p.add_run(f" — {ev.get('url')}")
            doc.add_heading(decision_ref["type"], level=3)
            if decision_ref["has_authoritative_reference"]:
                if decision_ref["fact"]:
                    p = doc.add_paragraph()
                    p.add_run("Project fact used: ").bold = True
                    p.add_run(decision_ref["fact"])
                if decision_ref["rule"]:
                    p = doc.add_paragraph()
                    p.add_run("Rule/source finding: ").bold = True
                    p.add_run(decision_ref["rule"])
                if decision_ref["conclusion"]:
                    p = doc.add_paragraph()
                    p.add_run("Conclusion: ").bold = True
                    p.add_run(decision_ref["conclusion"])
                doc.add_paragraph("Supporting source(s):")
                for eid in decision_ref["evidence_ids"]:
                    ev = ev_dict[eid]
                    p = doc.add_paragraph(style="List Bullet")
                    p.add_run(ev.get("title", eid))
                    if ev.get("url"):
                        p.add_run(f" — {ev.get('url')}")
            else:
                if decision_ref["applicability_evidence_ids"]:
                    doc.add_paragraph("The cited reference supports code applicability only; it is not treated as proof of a permit requirement or pathway.")
                    doc.add_paragraph("Use the cited applicability source as the starting point for the unresolved permit/pathway research.")
                    for eid in decision_ref["evidence_ids"]:
                        ev = ev_dict[eid]
                        p = doc.add_paragraph(style="List Bullet")
                        p.add_run(ev.get("title", eid))
                        if ev.get("url"):
                            p.add_run(f" — {ev.get('url')}")
                else:
                    doc.add_paragraph("No proposition-specific authoritative evidence was established for this decision.")
            if decision_ref["rule"]:
                p = doc.add_paragraph()
                p.add_run("Source finding: ").bold = True
                p.add_run(decision_ref["rule"])
            else:
                doc.add_paragraph("No proposition-specific authoritative rule was established for this decision.")
            p = doc.add_paragraph()
            p.add_run("Project information: ").bold = True
            p.add_run(f"{fact_statement} [{_pretty_fact_source(fact_source)}]")
            p = doc.add_paragraph()
            p.add_run("Permit: ").bold = True
            p.add_run(_pretty_status(item.get("permit")))
            doc.add_paragraph(str(item.get("permit_finding", "N/A")))
            p = doc.add_paragraph()
            p.add_run("Review / pathway: ").bold = True
            p.add_run(_pretty_pathway_status(item.get("pathway")))
            doc.add_paragraph(str(item.get("pathway_finding", "N/A")))
            questions = _discipline_questions(item)
            if questions:
                doc.add_paragraph("Questions to ask / research next:")
                for q in questions:
                    doc.add_paragraph(q, style="List Bullet")
            potential_issues = item.get("potential_issues") or []
            if potential_issues:
                doc.add_paragraph("Worth checking — AI-generated possible lead:")
                for lead in potential_issues:
                    if isinstance(lead, dict):
                        issue = str(lead.get("issue", "")).strip()
                        why = str(lead.get("why", "")).strip()
                    else:
                        issue, why = str(lead).strip(), ""
                    if issue:
                        doc.add_paragraph(issue, style="List Bullet")
                    if why:
                        p = doc.add_paragraph()
                        p.add_run("Basis: ").bold = True
                        p.add_run(why)
            missing = item.get("missing", [])
            if isinstance(missing, str):
                missing = [missing]
            if missing:
                doc.add_paragraph("Missing information:")
                for value in missing:
                    doc.add_paragraph(str(value), style="List Bullet")
            reopen = item.get("reopen", [])
            if isinstance(reopen, str):
                reopen = [reopen]
            if reopen:
                doc.add_paragraph("Reopen if:")
                for value in reopen:
                    doc.add_paragraph(str(value), style="List Bullet")
            doc.add_paragraph("")
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button(
            "📄 Download Word Report",
            data=buf.getvalue(),
            file_name="AHJ_Dossier.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with col2:
        json_data = json.dumps(
            {"project": {"state": state, "address": address, "date": str(project_date)}, "dossier": data},
            indent=2,
        )
        st.download_button(
            "💾 Save JSON Session",
            data=json_data,
            file_name="AHJ_Dossier.json",
            mime="application/json",
            use_container_width=True,
        )
