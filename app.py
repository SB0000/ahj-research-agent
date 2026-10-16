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

st.set_page_config(page_title="AHJ Research Assistant v27.23", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v27.23_logic_quality"

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
    return (evidence.get("proposition_type") == "THRESHOLD"
            and not evidence_proposition_integrity_errors(evidence)
            and not evidence_source_specificity_errors([evidence]))

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
        and not evidence_source_specificity_errors([evidence])
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
    if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT", "THRESHOLD"}:
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
        return [f"{eid}: missing rule proposition."]
    if ptype == "PERMIT_REQUIREMENT":
        if has_definitive_negative_permit_claim(rule) or re.search(
            r"\b(?:may|might|could|possibly|potentially)\b|\bwhether\b|\bnot necessarily\b", rule
        ):
            errors.append(f"{eid}: requirement is negative, uncertain, or conditional; cannot establish an affirmative permit requirement.")

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
            "PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT", "THRESHOLD"
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
    address_tokens = re.findall(r"[a-z0-9]+", addr)
    # Require the street number and following street-name token in the source
    # extract, not just a generic statement about unincorporated parcels.
    street = next((address_tokens[i:i+2] for i, token in enumerate(address_tokens[:-1]) if token.isdigit()), [])
    excerpt = _norm_text(evidence.get("source_excerpt"))
    site_linked = bool(street and re.search(r"\b" + re.escape(street[0]) + r"\s+" + re.escape(street[1]) + r"\b", excerpt))
    return site_linked and any(term in excerpt for term in boundary_terms)

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

def _jurisdiction_boundary_evidence_is_strong(evidence, address=""):
    """Require evidence that actually resolves the parcel boundary."""
    if not isinstance(evidence, dict):
        return False
    if str(evidence.get("proposition_type") or "").upper() != "JURISDICTION":
        return False
    if jurisdiction_evidence_integrity_errors(evidence) or evidence_proposition_integrity_errors(evidence):
        return False
    url = str(evidence.get("url") or "").strip()
    if not url or _is_grounding_redirect_url(url):
        return False
    try:
        from urllib.parse import urlparse
        path = (urlparse(url).path or "/").rstrip("/").lower() or "/"
    except Exception:
        return False
    if path in {"/", "/index.html", "/home", "/services", "/building", "/building-and-safety", "/planning", "/planning-and-zoning", "/zoning", "/permits", "/permit", "/development-services", "/development", "/public-works", "/engineering"}:
        return False
    text = _norm_text(" ".join([str(evidence.get("title") or ""), str(evidence.get("rule") or ""), str(evidence.get("retrieval_note") or "")]))
    lookup_signal = re.compile(r"\b(?:parcel|apn|assessor|property\s+record|property\s+lookup|gis|map|boundary|jurisdiction\s+lookup|municipal\s+limits?|city\s+limits?|incorporated|unincorporated)\b", re.I)
    return bool(lookup_signal.search(text)) and _jurisdiction_evidence_is_site_specific(evidence, address)

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
    strong = [e for e in site_specific if _jurisdiction_boundary_evidence_is_strong(e, address)]
    if strong:
        combined = _norm_text(" ".join(str(e.get("rule") or "") for e in strong))
        if "unincorporated" in combined:
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

def sanitize_unsupported_applicability_threshold_claims(data):
    """Remove numeric/categorical applicability triggers unless threshold evidence supports them."""
    if not isinstance(data, dict):
        return data
    evidence_by_id = {
        str(e.get("id")): e for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    threshold_claim = re.compile(
        r"\b(?:up to|within|under|over|exceed(?:s|ing)?|threshold|limit|valuation|20\s*%|[0-9]+\s*%)\b",
        re.I,
    )
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        app = item.get("applicability") or {}
        if not isinstance(app, dict):
            continue
        rule = _norm_text(app.get("rule"))
        if not rule or not threshold_claim.search(rule):
            continue
        discipline = str(item.get("type") or "Unknown")
        linked = app.get("evidence") or []
        if isinstance(linked, str):
            linked = [linked]
        valid = any(
            evidence_supports_threshold(evidence_by_id.get(str(eid)), discipline)
            and not evidence_source_specificity_errors([evidence_by_id.get(str(eid))])
            for eid in linked
            if evidence_by_id.get(str(eid))
        )
        if valid:
            continue
        app["rule"] = (
            f"The {discipline.lower()} applicability topic may be relevant, but the specific "
            "threshold or valuation limit is not established by the current evidence."
        )
        app["determination"] = "cannot_determine"
        app["relationship"] = "conditional"
        app["missing"] = "Proposition-specific threshold or applicability source for the stated trigger."
        app["evidence"] = []
        app["validation_note"] = (
            "Applicability rule was neutralized because it contained a numeric or categorical "
            "trigger without discipline-matched threshold evidence."
        )
    return data


def sanitize_unsubstantiated_conditional_pathways(data):
    """Downgrade conditional pathway claims that contain a concrete routing/review assertion without evidence."""
    if not isinstance(data, dict):
        return data
    evidence_by_id = {str(e.get("id")): e for e in (data.get("evidence") or []) if isinstance(e, dict) and e.get("id")}
    claim_patterns = [
        r"\b(?:clearance|required|must|requires?|reviewed|processed|submitted|submit|application|portal|plan review)\b",
        r"\bprior to\b",
        r"\bunder (?:the )?(?:main )?(?:electrical|building|fire|mechanical|plumbing) permit\b",
        r"\bvia\s+(?:epic[- ]la|portal|online|over[- ]the[- ]counter)\b",
    ]
    epistemic = re.compile(r"\b(?:not established|cannot determine|unable to determine|depends on an unresolved fact|current evidence does not establish)\b", re.I)
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        finding = _norm_text(item.get("pathway_finding"))
        if not finding or epistemic.search(finding):
            continue
        if not any(re.search(p, finding, re.I) for p in claim_patterns):
            continue
        ids = item.get("pathway_evidence") or item.get("process_evidence") or []
        if isinstance(ids, str): ids = [ids]
        valid = []
        for eid in ids:
            ev = evidence_by_id.get(str(eid))
            if not ev: continue
            ptype = str(ev.get("proposition_type") or "").upper()
            if ptype not in {"PATHWAY", "REVIEW_REQUIREMENT"}: continue
            if evidence_proposition_integrity_errors(ev): continue
            # Both pathway and review evidence must come from a proposition-specific
            # source. A generic agency landing page cannot establish routing or review.
            if evidence_source_specificity_errors([ev]): continue
            valid.append(str(eid))
        if valid:
            continue
        item["pathway"] = "CONDITIONAL"
        item["pathway_basis"] = "NOT_ESTABLISHED"
        item["pathway_evidence"] = []
        item["pathway_finding"] = (
            f"The {str(item.get('type') or 'discipline').lower()} processing pathway is not established by current evidence. "
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
            "PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT", "THRESHOLD"
        }:
            # Use the exact same source-specificity contract as the validator.
            # This is intentionally broader than the rule-text integrity check:
            # a specific-looking URL with a weak rule, or a generic portal URL,
            # must not survive as proposition-specific evidence.
            source_errors = evidence_source_specificity_errors([evidence])

        if (integrity_errors or source_errors) and original in EVIDENCE_PROPOSITION_TYPES and original != "OTHER":
            evidence["quarantined_proposition_type"] = original
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

# v27.23: deterministic source retrieval and code-adoption verification.
# No Gemini call is made by this subsystem. Failures remain unresolved.
_CODE_PAGE_CACHE = {}


def _official_public_url(url):
    from urllib.parse import urlsplit
    import ipaddress
    import socket
    parsed = urlsplit(str(url or ""))
    host = (parsed.hostname or "").lower()
    if parsed.scheme != "https" or parsed.username or parsed.password or parsed.port not in {None, 443}:
        return False
    # .us alone does not establish government ownership. State-government
    # namespaces are accepted; other official domains require a future adapter.
    if not (host.endswith(".gov") or re.search(r"(?:^|\.)state\.[a-z]{2}\.us$", host)):
        return False
    try:
        addresses = socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
        return bool(addresses) and all(ipaddress.ip_address(a[4][0]).is_global for a in addresses)
    except (OSError, ValueError):
        return False


def _adoption_html_text(html):
    from html.parser import HTMLParser
    class Reader(HTMLParser):
        def __init__(self):
            super().__init__(); self.parts = []; self.skip = 0; self.links = []; self.anchor = None
        def handle_starttag(self, tag, attrs):
            if tag in {"script", "style", "noscript"}: self.skip += 1
            if self.skip: return
            if tag in {"p", "li", "div", "br", "h1", "h2", "h3", "h4", "blockquote", "tr"}: self.parts.append("\n")
            if tag == "a": self.anchor = [dict(attrs).get("href", ""), []]
        def handle_endtag(self, tag):
            if tag in {"script", "style", "noscript"}: self.skip = max(0, self.skip - 1); return
            if self.skip: return
            if tag in {"p", "li", "div", "h1", "h2", "h3", "h4", "blockquote", "tr"}: self.parts.append("\n")
            if tag == "a" and self.anchor:
                self.links.append((self.anchor[0], "".join(self.anchor[1]))); self.anchor = None
        def handle_data(self, value):
            if self.skip: return
            self.parts.append(value)
            if self.anchor: self.anchor[1].append(value)
    reader = Reader(); reader.feed(html)
    text = "\n".join(re.sub(r"\s+", " ", line).strip() for line in "".join(reader.parts).splitlines() if line.strip())
    return text, reader.links


def fetch_official_code_page(url):
    """Bounded HTML retrieval, including redirect checks and an expiring cache."""
    import urllib.request
    from urllib.parse import urljoin
    url = str(url or "").strip()
    cache = st.session_state.setdefault("official_code_page_cache_v27_16", {}) if "st" in globals() else _CODE_PAGE_CACHE
    now = time.time()
    cached = cache.get(url)
    if cached and now - cached[0] < (86400 if cached[1].get("status") == "ok" else 900):
        return copy.deepcopy(cached[1])
    result = {"status": "unavailable", "requested_url": url}
    try:
        if not _official_public_url(url): raise ValueError("URL is not an allowed public government HTTPS source.")
        class Redirect(urllib.request.HTTPRedirectHandler):
            max_redirections = 3
            def redirect_request(self, req, fp, code, msg, headers, newurl):
                if not _official_public_url(newurl): raise ValueError("Redirect to an unapproved source.")
                return super().redirect_request(req, fp, code, msg, headers, newurl)
        request = urllib.request.Request(url, headers={"User-Agent": "AHJ-Research-Assistant/27.16", "Accept": "text/html"})
        with urllib.request.build_opener(Redirect()).open(request, timeout=8) as response:
            if response.headers.get_content_type() not in {"text/html", "application/xhtml+xml"}:
                raise ValueError("Automatic verification currently supports HTML adoption pages; document needs review.")
            raw = response.read(2_000_001)
            if len(raw) > 2_000_000: raise ValueError("Source exceeds retrieval size limit.")
            html = raw.decode(response.headers.get_content_charset() or "utf-8", errors="replace")
            text, links = _adoption_html_text(html)
            from html import unescape
            title_match=re.search(r"<title[^>]*>(.*?)</title>",html,re.I|re.S)
            page_title=re.sub(r"\s+"," ",unescape(re.sub(r"<[^>]+>"," ",title_match.group(1)))).strip() if title_match else ""
            result = {"status": "ok", "requested_url": url, "url": response.geturl(), "text": text, "title": page_title,
                      "links": [(urljoin(response.geturl(), href), label) for href, label in links],
                      "sha256": hashlib.sha256(raw).hexdigest(), "fetched_at": datetime.now(timezone.utc).isoformat()}
    except Exception as exc:
        result["error"] = str(exc)[:250]
    if len(cache) >= 32: cache.pop(next(iter(cache)))
    cache[url] = (now, copy.deepcopy(result))
    return result


def _code_family(name):
    text = _norm_text(name)
    for family, terms in (
        ("electrical", ("electrical", "oesc", "nec")), ("plumbing", ("plumbing", "opsc")),
        ("mechanical", ("mechanical", "omsc", "hvac")), ("energy", ("energy", "oeesc")),
        ("residential", ("residential", "orsc")), ("existing", ("existing building",)),
        ("structural", ("structural", "ossc")), ("fire", ("fire",)), ("building", ("building",))):
        if any(re.search(r"\b" + re.escape(term) + r"\b", text) for term in terms): return family
    return ""


def _adoption_date(text, label):
    # Dates are extracted from the SAME code block, never from another discipline.
    months = r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    match = re.search(r"\b" + label + r"\b\s*:?\s*(" + months + r"\.?\s+\d{1,2},?\s+20\d{2}|20\d{2}-\d{2}-\d{2})", text, re.I)
    if not match: return None
    value = match.group(1).replace(".", "").replace(",", "")
    for fmt in ("%b %d %Y", "%B %d %Y", "%Y-%m-%d"):
        try: return datetime.strptime(value, fmt).date().isoformat()
        except ValueError: pass
    return None


def parse_adoption_records(page):
    """Parse year/code-title blocks with their own dates; no model-provided rules."""
    if page.get("status") != "ok": return []
    lines = page.get("text", "").splitlines()
    starts = []
    for idx, line in enumerate(lines):
        match = re.match(r"\s*((?:19|20)\d{2})\s+(.{3,160}?\b(?:Code|Standard)\b[^\n]*)", line, re.I)
        if match and _code_family(match.group(0)): starts.append((idx, match))
    records = []
    for pos, (idx, match) in enumerate(starts):
        end = starts[pos+1][0] if pos+1 < len(starts) else len(lines)
        block = "\n".join(lines[idx:min(end, idx+16)])
        # Split chapter schedules need chapter-specific records, not one whole-code status.
        if re.search(r"\bchapter(?:s)?\s+\d", block, re.I): continue
        effective = _adoption_date(block, "effective")
        mandatory = _adoption_date(block, "mandatory")
        if not effective: continue
        name = re.sub(r"\s+New adoption.*$", "", match.group(0), flags=re.I).strip()
        records.append({"name": name, "family": _code_family(name), "edition": int(match.group(1)),
                        "effective_date": effective, "mandatory_date": mandatory,
                        "excerpt": block, "url": page.get("url"), "fetched_at": page.get("fetched_at"),
                        "sha256": page.get("sha256")})
    return records


def verify_code_adoptions(data, as_of_date=None, fetcher=None):
    """Replace model code-currency assertions with records from retrieved page text."""
    if not isinstance(data, dict): return data
    as_of_date = as_of_date or date.today()
    fetcher = fetcher or fetch_official_code_page
    evidence = {str(e.get("id")): e for e in data.get("evidence", []) if isinstance(e, dict)}
    audits = data.setdefault("_adoption_source_audit", {})
    # Repeated deterministic passes reuse the immutable fetch result from this run.
    def get_page(url):
        if url not in audits:
            audits[url] = fetcher(url) if len(audits) < 6 else {"status": "unavailable", "error": "Per-run source limit reached."}
        return audits[url]
    all_urls = set()
    for code in data.get("codes", []) or []:
        if not isinstance(code,dict): continue
        ids = code.get("evidence") or []
        if isinstance(ids,str): ids=[ids]
        for eid in ids:
            ev=evidence.get(eid) or {}
            url=str(ev.get("url") or "")
            if ev.get("proposition_type")=="CODE_CURRENCY" and url and not re.search(r"history|archive",url,re.I): all_urls.add(url)
    # Stable ordering and the same six-source budget; no additional AI call.
    for url in sorted(all_urls,key=lambda u:(not bool(re.search(r"adopted.codes|current.codes",u,re.I)),u)):
        page=get_page(url)
        if page.get("status")=="ok" and not parse_adoption_records(page):
            from urllib.parse import urlsplit
            link=next((u for u,label in page.get("links",[]) if re.search(r"adopted codes|current codes",label,re.I)
                and urlsplit(u).hostname==urlsplit(url).hostname and u!=url),None)
            if link:get_page(link)
    for code in data.get("codes", []) or []:
        if not isinstance(code, dict): continue
        family = _code_family(code.get("name"))
        original = str(code.get("name") or "")
        ids = code.get("evidence") or []
        if isinstance(ids, str): ids = [ids]
        urls = list(dict.fromkeys(str(evidence[eid].get("url") or "") for eid in ids
                    if eid in evidence and evidence[eid].get("proposition_type") == "CODE_CURRENCY"))
        if len(_extract_edition_years(original)) > 1:
            code["status"] = "CONDITIONAL"
            code.pop("adoption_record", None)
            code["validation_note"] = "Local adoption and state base edition are combined in this entry. Verify them as separate records; do not substitute one edition year for the other."
            continue
        records = [record for page in audits.values() for record in parse_adoption_records(page)]
        for url in urls[:3]:
            if not url or re.search(r"history|archive", url, re.I): continue
            page = get_page(url)
            records.extend(parse_adoption_records(page))
            # A source may be an index; follow ONE explicitly named adoption link.
            if page.get("status") == "ok" and not parse_adoption_records(page):
                from urllib.parse import urlsplit
                link = next((u for u, label in page.get("links", [])
                    if re.search(r"adopted codes|current codes", label, re.I)
                    and urlsplit(u).hostname == urlsplit(url).hostname and u != url), None)
                if link: records.extend(parse_adoption_records(get_page(link)))
        state = str((data.get("_research_context") or {}).get("state") or "").strip()
        family_records = [r for r in records if family and r["family"] == family
                          and (not state or state.lower() in r["name"].lower())]
        eligible = [r for r in family_records if date.fromisoformat(r["effective_date"]) <= as_of_date
                    and (not r["mandatory_date"] or date.fromisoformat(r["mandatory_date"]) <= as_of_date)]
        if eligible:
            newest = max(r["effective_date"] for r in eligible)
            choices = [r for r in eligible if r["effective_date"] == newest]
            if len({r["edition"] for r in choices}) == 1:
                record = choices[0]
                code["generated_name"] = code.get("generated_name") or original
                code["name"] = re.sub(r"\b(?:19|20)\d{2}\b", str(record["edition"]), original, count=1) if _extract_edition_years(original) else record["name"]
                code["status"] = "CURRENT"
                source_ids = [eid for eid, ev in evidence.items() if str(ev.get("url") or "") == record["url"] or (audits.get(str(ev.get("url") or "")) or {}).get("url") == record["url"]]
                if source_ids: code["evidence"] = source_ids
                code["adoption_record"] = record
                code["validation_note"] = "Edition and dates checked against retrieved official page text; project-specific application is separate."
                continue
        code["status"] = "CONDITIONAL"
        code.pop("adoption_record", None)
        code["validation_note"] = "Code currency not verified from a dated, family-specific official HTML adoption record. Review the source, transition dates, or document manually."
    # Keep the evidence appendix aligned with the independently checked code
    # list. Retain Gemini's wording privately for comparison, never as proof.
    checked_by_url = {}
    for code in data.get("codes", []) or []:
        record = code.get("adoption_record") if isinstance(code, dict) else None
        if record:
            checked_by_url.setdefault(record["url"], []).append(record)
    for ev in evidence.values():
        if ev.get("proposition_type") != "CODE_CURRENCY": continue
        requested = str(ev.get("url") or "")
        page = audits.get(requested) or {}
        if page.get("title"):
            ev.setdefault("generated_title",ev.get("title"))
            ev["title"]=page["title"]
        records = checked_by_url.get(page.get("url")) or checked_by_url.get(requested) or []
        if not records:
            candidates=[r for r in parse_adoption_records(page)
                if date.fromisoformat(r["effective_date"])<=as_of_date
                and (not r["mandatory_date"] or date.fromisoformat(r["mandatory_date"])<=as_of_date)]
            for family in sorted({r["family"] for r in candidates}):
                group=[r for r in candidates if r["family"]==family]
                newest=max(r["effective_date"] for r in group)
                choices=[r for r in group if r["effective_date"]==newest]
                if len({r["edition"] for r in choices})==1:records.append(choices[0])
        if not records:
            ev["source_verification"]="UNVERIFIED_EXTRACTION"
            ev["validation_note"]="Generated adoption summary not matched to a dated official source record; do not treat it as verified code currency."
            continue
        ev["source_verification"]="SOURCE_CHECKED"
        ev.setdefault("generated_rule", ev.get("rule"))
        ev["rule"] = "; ".join(dict.fromkeys(
            r["name"] + " effective " + r["effective_date"] +
            (", mandatory " + r["mandatory_date"] if r.get("mandatory_date") else "") for r in records))
        ev["source_excerpt"] = "\n\n".join(dict.fromkeys(r["excerpt"] for r in records))
        ev["validation_note"] = "Code adoption details extracted from retrieved official page text."
    return data


def code_currency_integrity_errors(code, evidence_by_id, as_of_date=None):
    if not isinstance(code, dict) or code.get("status") != "CURRENT": return []
    record = code.get("adoption_record") or {}
    if not record.get("sha256") or not record.get("url") or not record.get("excerpt"):
        return ["Current code lacks a retrieved, dated official adoption record."]
    if _extract_edition_years(code.get("name")) != {record.get("edition")}:
        return ["Code edition conflicts with the retrieved adoption record."]
    if _code_family(code.get("name")) != record.get("family"):
        return ["Code family conflicts with the retrieved adoption record."]
    as_of_date = as_of_date or date.today()
    try:
        if date.fromisoformat(record["effective_date"]) > as_of_date:
            return ["Code is not effective as of the project date."]
        if record.get("mandatory_date") and date.fromisoformat(record["mandatory_date"]) > as_of_date:
            return ["Code remains within a transition period; review the applicable edition."]
    except (KeyError, ValueError): return ["Code adoption dates are not verified."]
    return []

def sanitize_invalid_current_codes(data, as_of_date=None):
    return verify_code_adoptions(data, as_of_date)

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
    if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT", "THRESHOLD"}:
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
        if best and _norm_text(best.get("title")) == title:
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
        if ptype not in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "ENTITLEMENT", "REVIEW_REQUIREMENT", "THRESHOLD"}:
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
        if parsed.scheme not in {"http", "https"} or not parsed.hostname:
            errors.append(f"{ev.get('id', 'Evidence')}: source must be an absolute HTTP(S) URL.")
            continue
        if ptype == "THRESHOLD":
            check = ev.get("_threshold_source_check") or {}
            if check.get("status") != "verified":
                errors.append(f"{ev.get('id', 'Evidence')}: threshold proposition was not matched to retrieved official source text.")
            if re.search(r"/(?:building-services|building-and-safety|design-criteria|planning-and-development-services)/?$", path, re.I):
                errors.append(f"{ev.get('id', 'Evidence')}: generic department/design page cannot establish this threshold.")
        excerpt = str(ev.get("source_excerpt") or "").strip()
        if not excerpt:
            errors.append(f"{ev.get('id', 'Evidence')}: missing source excerpt for proposition-specific review.")
        if ptype == "PATHWAY" and re.search(r"/(?:payments?-and-fees|fees|payments?)/?$", path, re.I):
            errors.append(f"{ev.get('id', 'Evidence')}: a payments/fees page cannot establish the submission route without a reviewed source.")
        # Agency permit hubs often have a path such as /building-and-safety/permits.
        # The word "permit" in a URL is NOT enough to make the page proposition-specific.
        # Only a deeper permit-specific document/page may support a definitive permit rule.
        generic_permit_hub = bool(re.fullmatch(
            r"/(?:building-and-safety/)?(?:permits?|permit-center|permitcenter|applications?|applications-and-forms)",
            path, re.I,
        ))
        if path in generic_paths or generic_permit_hub or _is_generic_regulatory_landing_page(ev):
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
        if not any(_jurisdiction_evidence_is_site_specific(evidence_by_id[eid], str((data.get("_research_context") or {}).get("address") or data.get("address") or "")) for eid in valid_jurisdiction_ids):
            errors.append("Verified jurisdiction requires site-specific parcel/boundary evidence; postal city or generic agency coverage is insufficient.")

    for code in data.get("codes", []):
        for eid in code.get("evidence", []):
            if eid not in evidence_ids: errors.append(f"Code '{code.get('name')}' references nonexistent evidence: {eid}")
        if code.get("status") == "CURRENT" and not code.get("evidence"):
            errors.append(f"Current code '{code.get('name')}' must have evidence.")
        errors.extend(code_currency_integrity_errors(code, evidence_by_id))

    bottom_line_evidence = data.get("bottom_line_evidence") or []
    for evidence_id in bottom_line_evidence:
        if evidence_id not in evidence_by_id:
            errors.append(f"Bottom Line: unknown evidence ID '{evidence_id}'.")
    if not bottom_line_evidence:
        expected = build_validated_bottom_line(copy.deepcopy(data)).get("bottom_line", "")
        asserted = ((data.get("jurisdiction") or {}).get("status") == "VERIFIED" or any(
            isinstance(i, dict) and i.get("permit") in {"VERIFIED_REQUIRED", "NOT_APPLICABLE"}
            for i in data.get("disciplines", []) or []))
        # Evidence is required for established legal outcomes. The application's
        # exact, non-establishment summary does not manufacture a proposition.
        if asserted or str(data.get("bottom_line") or "") != expected:
            errors.append("Bottom Line: established or free-form regulatory conclusions require supporting evidence.")

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

def _response_value(obj, snake, camel=None, default=None):
    if isinstance(obj, dict): return obj.get(snake, obj.get(camel, default) if camel else default)
    return getattr(obj, snake, default)


def inspect_generation_response(response):
    """Capture final-answer text and stop diagnostics before deciding acceptability.

    Thought text/signatures are deliberately not stored as report content.
    """
    candidates = _response_value(response, "candidates", default=[]) or []
    info = {"candidate_count":len(candidates), "response_id":_response_value(response,"response_id","responseId"),
            "model_version":_response_value(response,"model_version","modelVersion")}
    meta = _response_value(response,"usage_metadata","usageMetadata")
    if hasattr(meta,"model_dump"):
        try: info["usage_metadata_raw"] = meta.model_dump(mode="json",exclude_none=False)
        except Exception: pass
    if "usage_metadata_raw" not in info:
        info["usage_metadata_raw"] = {k:_response_value(meta,k,c) for k,c in (
            ("prompt_token_count","promptTokenCount"),("candidates_token_count","candidatesTokenCount"),
            ("thoughts_token_count","thoughtsTokenCount"),("tool_use_prompt_token_count","toolUsePromptTokenCount"),
            ("total_token_count","totalTokenCount"),("cached_content_token_count","cachedContentTokenCount"))}
    if not candidates:
        info["finish_reason_normalized"] = "NO_CANDIDATE"
        return "",info
    candidate = candidates[0]
    reason = _response_value(candidate,"finish_reason","finishReason","")
    reason_name = getattr(reason,"name",None) or str(reason or "").rsplit(".",1)[-1]
    info.update(finish_reason=str(reason),finish_reason_normalized=reason_name,
        finish_message=_response_value(candidate,"finish_message","finishMessage"),
        candidate_token_count=_response_value(candidate,"token_count","tokenCount"),
        safety_ratings=[str(r) for r in (_response_value(candidate,"safety_ratings","safetyRatings",[]) or [])])
    content = _response_value(candidate,"content")
    parts = _response_value(content,"parts",default=[]) or []
    text_parts = []; descriptors=[]
    for part in parts:
        text = _response_value(part,"text",default="") or ""
        thought = _response_value(part,"thought",default=False) is True
        descriptors.append({"thought":thought,"text_characters":len(text),
            "function_call":bool(_response_value(part,"function_call","functionCall"))})
        if isinstance(text,str) and text and not thought: text_parts.append(text)
    # Some test doubles / SDK versions expose only the text convenience property.
    # Do not use that fallback if explicit thought-only parts were returned.
    if not parts:
        try: final_text = _response_value(response,"text",default="") or ""
        except Exception: final_text = ""
    else: final_text = "".join(text_parts)
    info["part_descriptors"] = descriptors
    info["raw_response_text"] = final_text
    info["text_length"] = len(final_text)
    return final_text, info


def normalize_compact_research(data, scope):
    """Supply structural defaults locally, never missing legal conclusions."""
    if not isinstance(data,dict): raise ValueError("Research JSON must be an object.")
    for field in ("evidence","codes","disciplines"):
        if field not in data: data[field]=[]
        if not isinstance(data[field],list): raise ValueError(field + " must be an array.")
    data.setdefault("bottom_line",""); data.setdefault("bottom_line_evidence",[])
    data.setdefault("research_completeness",{"status":"PARTIAL","reason":"Further AHJ or project-record research may be needed.","critical_missing":[]})
    data.setdefault("jurisdiction",{"status":"CONDITIONAL","county":"Not established","city":"Not yet established","ahj":"Unconfirmed","evidence":[]})
    for code in data["codes"]:
        if isinstance(code,dict): code.setdefault("status","CONDITIONAL"); code.setdefault("evidence",[])
    for item in data["disciplines"]:
        if not isinstance(item,dict): continue
        quotes = _scope_excerpts(scope,item.get("type"))
        item.setdefault("scope_likelihood","LIKELY" if quotes else "POSSIBLE")
        item.setdefault("scope_basis","")
        app = item.setdefault("applicability",{})
        if not isinstance(app,dict): continue
        app.setdefault("rule",""); app.setdefault("evidence",[])
        app.setdefault("determination","cannot_determine"); app.setdefault("relationship","not_established")
        app.setdefault("missing","Discipline-specific applicability support")
        app.setdefault("fact",{"statement":quotes[0] if quotes else "Relevant project facts still need confirmation.","source":"USER_PROVIDED" if quotes else "UNKNOWN"})
        for kind in ("permit","pathway"):
            item.setdefault(kind,"UNKNOWN"); item.setdefault(kind+"_basis","NOT_ESTABLISHED")
            item.setdefault(kind+"_evidence",[]); item.setdefault(kind+"_finding","")
        for field in ("authority_evidence","missing","reopen","reported_guidance","decision_followups"):
            item.setdefault(field,[])
        item.setdefault("process_confidence","NOT_ESTABLISHED");item.setdefault("expected_process","")
    return data


def record_gemini_usage(model, response, caller="", prompt_hash="", elapsed_seconds=None):
    """Record one actual Gemini API response for per-request cost analysis."""
    meta = _response_value(response, "usage_metadata", "usageMetadata")
    now = datetime.now(timezone.utc).isoformat()
    entry = {
        "timestamp_utc": now,
        "caller": caller,
        "model": model,
        "prompt_hash": prompt_hash,
        "elapsed_seconds": round(float(elapsed_seconds), 3) if elapsed_seconds is not None else None,
        "input_tokens": _response_value(meta, "prompt_token_count", "promptTokenCount"),
        "output_tokens": _response_value(meta, "candidates_token_count", "candidatesTokenCount"),
        "thoughts_tokens": _response_value(meta, "thoughts_token_count", "thoughtsTokenCount"),
        "total_tokens": _response_value(meta, "total_token_count", "totalTokenCount"),
        "cached_tokens": _response_value(meta, "cached_content_token_count", "cachedContentTokenCount"),
        "tool_prompt_tokens": _response_value(meta, "tool_use_prompt_token_count", "toolUsePromptTokenCount"),
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
RESEARCH_DECISION_CONTRACT = "No model-supplied field can override the deterministic evidence contract. Never invent facts, sources, quotations, legal thresholds or requirements."

@cache_success_only(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": 1,
        "request_config":{"model":"gemini-3.6-flash","max_output_tokens":32768,"thinking_level":"medium","google_search":True},
        "project_context":{"address":str(address),"state":str(state),"project_date":str(project_date),
            "scope":str(sow_text),"researcher_notes":str(globals().get("researcher_notes", "")),"entitlements":str(globals().get("existing_permit", ""))}}
    try:
        try:
            from importlib.metadata import version as package_version
            debug_info["google_genai_version"] = package_version("google-genai")
        except Exception:
            debug_info["google_genai_version"] = "unavailable"
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_level="medium"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        _api_started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=prompt_text + "\n" + RESEARCH_DECISION_CONTRACT,
            config=config,
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="initial_research",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - _api_started
        )
        text, response_debug = inspect_generation_response(response)
        debug_info.update(response_debug)
        finish_reason = response_debug.get("finish_reason_normalized")
        if finish_reason not in {"STOP","MAX_TOKENS"}:
            debug_info["status"] = "failed"
            debug_info["error_type"] = "No Candidates" if finish_reason == "NO_CANDIDATE" else "Early Stop"
            return {"data":None,"error":True,"retry":False,"msg":f"Gemini returned {finish_reason or 'an unspecified stop reason'}. No automatic retry was made.","debug":debug_info}
        if not text:
            debug_info["status"] = "failed"
            debug_info["error_type"] = "MAX_TOKENS_NO_TEXT" if finish_reason == "MAX_TOKENS" else "Empty Text"
            message = ("Gemini reported MAX_TOKENS but returned no final-answer text. The reported token counts do not establish why it stopped. No automatic retry was made."
                if finish_reason == "MAX_TOKENS" else "Gemini returned no final-answer text. No automatic retry was made.")
            return {"data":None,"error":True,"retry":False,"msg":message,"debug":debug_info}
        try:
            # Never close braces, invent missing fields from a cut-off object,
            # or recover isolated nested objects from a truncated response.
            if finish_reason == "MAX_TOKENS":
                candidate_json = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
                data = json.loads(candidate_json)
            else:
                data = extract_json(text)
            if not isinstance(data,dict): raise ValueError("Research response must be one JSON object.")
            if isinstance(data, dict):
                data.pop("_adoption_source_audit", None)
                data.pop("_threshold_source_audit", None)
                for ev in data.get("evidence", []) or []:
                    if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
                for code in data.get("codes", []) or []:
                    if isinstance(code, dict): code.pop("adoption_record", None)
            if isinstance(data, dict):
                st.session_state["last_raw_dossier"] = copy.deepcopy(data)
                debug_info["raw_dossier"] = copy.deepcopy(data)
                data["_research_context"] = {"scope": str(sow_text), "address": str(address), "researcher_notes": str(globals().get("researcher_notes", ""))}
            data = normalize_compact_research(data, str(sow_text))
            if finish_reason == "MAX_TOKENS":
                data["_generation_incomplete"] = True
                data["research_completeness"] = {"status":"PARTIAL","reason":"Provider ended generation early; only the complete JSON object returned is being evaluated. Further research may remain.","critical_missing":[]}
                debug_info["response_warning"] = "Provider reported MAX_TOKENS. Complete JSON was returned; findings still require the normal evidence checks."
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["status"] = "failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "error": True, "retry": False, "msg": "Gemini returned incomplete or invalid research JSON. The available final-answer text was saved in the snapshot. No automatic retry was made.", "debug": debug_info}
        
        debug_info["status"] = "partial" if finish_reason == "MAX_TOKENS" else "success"
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
            model="gemini-3.6-flash", contents=recovery_prompt + "\n" + RESEARCH_DECISION_CONTRACT, config=config
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
            if isinstance(data, dict):
                data.pop("_adoption_source_audit", None)
                data.pop("_threshold_source_audit", None)
                for ev in data.get("evidence", []) or []:
                    if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
                for code in data.get("codes", []) or []:
                    if isinstance(code, dict): code.pop("adoption_record", None)
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
            model="gemini-3.6-flash", contents=recovery_prompt + "\n" + RESEARCH_DECISION_CONTRACT, config=config
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
            if isinstance(data, dict):
                data.pop("_adoption_source_audit", None)
                data.pop("_threshold_source_audit", None)
                for ev in data.get("evidence", []) or []:
                    if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
                for code in data.get("codes", []) or []:
                    if isinstance(code, dict): code.pop("adoption_record", None)
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
            model="gemini-3.6-flash", contents=recovery_prompt + "\n" + RESEARCH_DECISION_CONTRACT, config=config
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
            if isinstance(data, dict):
                data.pop("_adoption_source_audit", None)
                data.pop("_threshold_source_audit", None)
                for ev in data.get("evidence", []) or []:
                    if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
                for code in data.get("codes", []) or []:
                    if isinstance(code, dict): code.pop("adoption_record", None)
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
            contents=retry_prompt + "\n" + RESEARCH_DECISION_CONTRACT,
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
            if isinstance(data, dict):
                data.pop("_adoption_source_audit", None)
                data.pop("_threshold_source_audit", None)
                for ev in data.get("evidence", []) or []:
                    if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
                for code in data.get("codes", []) or []:
                    if isinstance(code, dict): code.pop("adoption_record", None)
            if isinstance(data, dict):
                st.session_state["last_raw_dossier"] = copy.deepcopy(data)
                debug_info["raw_dossier"] = copy.deepcopy(data)
                data["_research_context"] = {"scope": str(sow_text), "address": str(address)}
            data = _normalize_evidence_urls(data)
            data = _repair_grounding_redirect_urls(data, response)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "retry": False, "msg": "Retry produced invalid JSON.", "debug": debug_info}
        
        data, validation_errors = run_deterministic_contract_pass(
            data, address, project_date, state
        )
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
        "scope", "address", "state", "project_date", "_research_context", "_adoption_source_audit", "_threshold_source_audit"
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
            if isinstance(data, dict):
                data.pop("_adoption_source_audit", None)
                data.pop("_threshold_source_audit", None)
                for ev in data.get("evidence", []) or []:
                    if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
                for code in data.get("codes", []) or []:
                    if isinstance(code, dict): code.pop("adoption_record", None)
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

def _indefinite_article(value):
    """Return a/an for simple discipline labels used in deterministic prose."""
    text = str(value or "").strip().lower()
    return "an" if text[:1] in {"a", "e", "i", "o", "u"} else "a"

# ============================================================
# DETERMINISTIC PERMIT DECISION ENGINE
# ============================================================
def derive_permit_statuses_from_evidence(data):
    """Require both proposition evidence and an explicit scope match; preserve conflicts."""
    if not isinstance(data, dict):
        return data
    evidence_by_id = {str(e.get("id")): e for e in data.get("evidence", []) if isinstance(e, dict) and e.get("id")}
    scope = str((data.get("_research_context") or {}).get("scope") or "")
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown")
        ids = item.get("permit_evidence") or []
        if isinstance(ids, str): ids = [ids]
        ids = ids if isinstance(ids, list) else []
        # Orphans remain useful research candidates, but cannot establish an outcome
        # merely because their discipline label matches.
        candidates = list(dict.fromkeys(ids + list(evidence_by_id)))
        req = evidence_ids_supporting_type(candidates, evidence_by_id, "PERMIT_REQUIREMENT", discipline)
        exempt = evidence_ids_supporting_type(candidates, evidence_by_id, "PERMIT_EXEMPTION", discipline)
        item["permit_candidate_evidence"] = list(dict.fromkeys(req + exempt))
        match = item.get("permit_scope_match") or {}
        matched = []
        if isinstance(match, dict) and match.get("conditions_satisfied") is True:
            eid = str(match.get("evidence_id") or "")
            scope_quote = _norm_text(match.get("scope_quote"))
            rule_quote = _norm_text(match.get("rule_quote"))
            ev = evidence_by_id.get(eid) or {}
            if (eid in req + exempt and scope_quote and rule_quote and scope
                    and scope_quote in _norm_text(scope)
                    and rule_quote in _norm_text(ev.get("source_excerpt"))
                    and not match.get("missing_facts")):
                matched = [eid]
        conflict = bool(req and exempt)
        jurisdiction_blocked = str((data.get("jurisdiction") or {}).get("status") or "").upper() != "VERIFIED"
        if jurisdiction_blocked:
            matched = []
        item["permit_blocker"] = ("JURISDICTION" if jurisdiction_blocked else "EVIDENCE_CONFLICT" if conflict else "NONE" if matched else "SCOPE_MATCH" if req or exempt else "SOURCE")
        if jurisdiction_blocked:
            item["permit"] = "UNKNOWN"
            reason = "The permit decision is blocked by unverified parcel jurisdiction; missing equipment details are not the established blocker."
        elif conflict:
            item["permit"] = "UNKNOWN"
            reason = "Requirement and exemption evidence both exist; their scope conditions must be reconciled before deciding."
        elif matched:
            item["permit"] = "NOT_APPLICABLE" if matched[0] in exempt else "VERIFIED_REQUIRED"
            item["permit_basis"] = "DIRECT_EVIDENCE"
            item["permit_evidence"] = matched
            item["permit_finding"] = ("The cited exemption matches the stated scope." if matched[0] in exempt
                else "The cited permit requirement matches the stated scope.")
            continue
        else:
            item["permit"] = "UNKNOWN"
            reason = ("A permit rule was located, but its conditions have not been matched to the stated scope."
                      if req or exempt else "This research run did not establish the permit requirement. A direct source or agency confirmation is still needed.")
        item["permit_basis"] = "NOT_ESTABLISHED"
        item["permit_evidence"] = []
        item["permit_finding"] = reason
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
            item["process_finding"] = "No authoritative source found in this research run established the review/submission process for this scope."
            item["process_evidence"] = []
            for key in ("process_reviewer", "process_route", "process_relationship", "process_trigger"):
                item.pop(key, None)
    return data



# ============================================================
# v27.23 — ONE-SHOT TARGETED EVIDENCE RESOLUTION + LOGIC QUALITY
# ============================================================
# The initial research pass sometimes identifies a highly relevant permit rule
# but attaches it to a generic agency landing page. The evidence firewall is
# correct to quarantine that item. Rather than loosening the firewall or running
# broad permit-recovery loops, v27.13 performs at most ONE small Google-grounded
# call to locate direct authoritative sources for a few high-value quarantined
# permit propositions.

_EVIDENCE_RESOLUTION_MAX_TARGETS = 5


def _candidate_family_from_text(text):
    text = _norm_text(text)
    checks = (
        ("electrical", ("electrical", "panel", "branch circuit", "service amperage", "lighting system")),
        ("plumbing", ("plumbing", "water heater", "sewer lateral", "water supply", "fixture")),
        ("mechanical", ("mechanical", "hvac", "ductwork", "heating", "exhaust system")),
        ("structural", ("structural", "foundation", "anchorage", "framing")),
        ("fire", ("fire alarm", "fire safety", "fire department", "smoke detector")),
        ("building", ("building permit", "building code", "alter", "repair", "demolish")),
    )
    for family, terms in checks:
        if any(term in text for term in terms):
            return family
    return ""


def _resolution_family_for_discipline(name):
    text = _norm_text(name)
    if "plumb" in text:
        return "plumbing"
    if "mechan" in text or "hvac" in text:
        return "mechanical"
    if "electric" in text:
        return "electrical"
    if "struct" in text:
        return "structural"
    if "fire" in text:
        return "fire"
    if "building" in text or "architect" in text:
        return "building"
    return discipline_family(name)


def quarantined_permit_resolution_targets(data):
    """Return a short list of unresolved disciplines with a quarantined permit claim.

    This is intentionally narrow. It does not search every unresolved discipline.
    A target exists only when the initial pass already produced a permit-specific
    proposition, but that proposition could not be used because its source was
    generic/OTHER. That keeps the second call focused and inexpensive.
    """
    if not isinstance(data, dict):
        return []
    unresolved = {}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        if str(item.get("permit") or "UNKNOWN").upper() not in {"UNKNOWN", "CONDITIONAL", "INFERRED"}:
            continue
        name = str(item.get("type") or "").strip()
        if name:
            unresolved.setdefault(_resolution_family_for_discipline(name), []).append(name)

    targets = []
    seen = set()
    permit_claim = re.compile(
        r"\b(?:permit|approval|license)\b[^.!?;:]{0,220}\b(?:required|needed|necessary)\b|"
        r"\b(?:requires?|must obtain|shall obtain)\b[^.!?;:]{0,220}\b(?:permit|approval|license)\b",
        re.I,
    )
    for ev in data.get("evidence", []) or []:
        if not isinstance(ev, dict):
            continue
        # Only rescue evidence the deterministic firewall did NOT accept as a
        # permit proposition. Never re-research already-valid permit evidence.
        if str(ev.get("proposition_type") or "OTHER").upper() != "OTHER":
            continue
        rule = str(ev.get("rule") or "").strip()
        title = str(ev.get("title") or "").strip()
        if not rule or not permit_claim.search(rule):
            continue
        family = _candidate_family_from_text(f"{title} {rule}")
        names = unresolved.get(family) or []
        if not names:
            continue
        for discipline in names[:1]:
            key = (family, _norm_text(rule))
            if key in seen:
                continue
            seen.add(key)
            targets.append({
                "discipline": discipline,
                "candidate_title": title,
                "candidate_rule": rule,
                "rejected_url": str(ev.get("url") or "").strip(),
                "reason": "Initial permit proposition was quarantined because its source was not proposition-specific.",
            })
            break
        if len(targets) >= _EVIDENCE_RESOLUTION_MAX_TARGETS:
            break
    return targets


def build_targeted_evidence_resolution_prompt(data, targets, address, state, project_date, sow_text):
    jurisdiction = data.get("jurisdiction") or {}
    return f"""
You are performing ONE narrow evidence-resolution pass for an AHJ research dossier.
Return ONLY valid JSON. Use Google Search. Do not regenerate the dossier.

PROJECT
Address: {address}
State: {state}
Research date: {project_date}
Verified jurisdiction context: {json.dumps(jurisdiction, default=str)}
Scope of work: {sow_text}

TARGETS
{json.dumps(targets, indent=2)}

TASK
For each target, try to locate a DIRECT AUTHORITATIVE source that actually contains
that target's permit proposition for the verified AHJ. Prefer official code text,
ordinance text, a permit-specific agency page, checklist, application/instruction,
or other proposition-specific government document.

STRICT RULES
- Do not use the rejected generic landing-page URL as proof unless the proposition
  is visibly stated on that exact page.
- Do not invent section numbers, titles, URLs, or quotations.
- A human must be able to open the returned URL and find the stated rule.
- Do not infer a permit from code applicability, a threshold, plan review, or a
  general agency description.
- Do not return third-party summaries as authoritative evidence.
- If you cannot locate a direct authoritative source for a target, return
  resolved=false for that target. That is a successful outcome.
- Resolve only the listed targets. Do not add unrelated research.

RETURN SCHEMA
{{
  "resolutions": [
    {{
      "discipline": "exact target discipline",
      "resolved": true,
      "scope_match": {{"scope_quote": "verbatim supplied scope", "rule_quote": "verbatim source excerpt", "conditions_satisfied": false, "missing_facts": []}},
      "evidence": {{
        "title": "actual source title",
        "url": "direct authoritative URL",
        "authority": "issuing government authority",
        "jurisdiction_level": "state|county|city|federal|tribal|other",
        "discipline": "exact target discipline",
        "proposition_type": "PERMIT_REQUIREMENT",
        "source_type": "code|ordinance|permit_page|checklist|application|interpretation|other",
        "retrieval_note": "where on the source the proposition can be found",
        "source_excerpt": "short verbatim passage actually retrieved",
        "rule": "the specific permit proposition supported by the source"
      }}
    }}
  ]
}}
"""


@cache_success_only(ttl=3600)
def cached_gemini_targeted_evidence_resolution(prompt_hash, prompt_text):
    time.sleep(0.5)
    debug_info = {"status": "processing", "attempt": "targeted_evidence_resolution"}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=8192,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        started = time.monotonic()
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=prompt_text + "\n" + RESEARCH_DECISION_CONTRACT,
            config=config,
        )
        debug_info["api_usage"] = record_gemini_usage(
            "gemini-3.6-flash", response, caller="targeted_evidence_resolution",
            prompt_hash=prompt_hash, elapsed_seconds=time.monotonic() - started,
        )
        text = getattr(response, "text", None)
        if not text:
            try:
                text = response.candidates[0].content.parts[0].text
            except Exception:
                text = None
        if not text:
            return {"data": None, "error": True, "msg": "Evidence-resolution call returned no text.", "debug": debug_info}
        payload = extract_json(text)
        # Reuse grounding redirect repair on the nested evidence objects.
        temp = {"evidence": []}
        for row in payload.get("resolutions", []) if isinstance(payload, dict) else []:
            if isinstance(row, dict) and isinstance(row.get("evidence"), dict):
                temp["evidence"].append(row["evidence"])
        temp = _normalize_evidence_urls(temp)
        temp = _repair_grounding_redirect_urls(temp, response)
        repaired_iter = iter(temp.get("evidence", []))
        for row in payload.get("resolutions", []) if isinstance(payload, dict) else []:
            if isinstance(row, dict) and isinstance(row.get("evidence"), dict):
                row["evidence"] = next(repaired_iter, row["evidence"])
        debug_info["status"] = "success"
        return {"data": payload, "error": False, "debug": debug_info}
    except Exception as exc:
        debug_info["api_usage_error"] = record_gemini_exception(
            "gemini-3.6-flash", caller="targeted_evidence_resolution",
            prompt_hash=prompt_hash,
            elapsed_seconds=(time.monotonic() - started) if "started" in locals() else None,
            error=str(exc),
        )
        debug_info["exception"] = str(exc)
        return {"data": None, "error": True, "msg": f"Evidence-resolution error: {str(exc)[:250]}", "debug": debug_info}


def merge_targeted_evidence_resolution(data, payload, targets):
    """Merge only direct permit evidence for the exact requested disciplines.

    The normal deterministic contract runs after this merge. Therefore a bad or
    still-generic source is simply quarantined again rather than being trusted.
    """
    if not isinstance(data, dict) or not isinstance(payload, dict):
        return data
    allowed = {str(t.get("discipline") or "").strip() for t in targets if isinstance(t, dict)}
    evidence = data.setdefault("evidence", [])
    existing_ids = {str(ev.get("id")) for ev in evidence if isinstance(ev, dict) and ev.get("id")}
    next_num = 1

    def new_id():
        nonlocal next_num
        while f"R{next_num}" in existing_ids:
            next_num += 1
        value = f"R{next_num}"
        existing_ids.add(value)
        next_num += 1
        return value

    by_name = {str(i.get("type") or "").strip(): i for i in data.get("disciplines", []) if isinstance(i, dict)}
    merged = []
    for row in payload.get("resolutions", []) or []:
        if not isinstance(row, dict) or row.get("resolved") is not True:
            continue
        discipline = str(row.get("discipline") or "").strip()
        ev = row.get("evidence")
        if discipline not in allowed or discipline not in by_name or not isinstance(ev, dict):
            continue
        ev = copy.deepcopy(ev)
        ev["id"] = new_id()
        ev["discipline"] = discipline
        ev["proposition_type"] = "PERMIT_REQUIREMENT"
        # Preflight the same proposition/source-specificity firewalls used by the dossier.
        if evidence_source_specificity_errors([ev]) or evidence_proposition_integrity_errors(ev):
            continue
        evidence.append(ev)
        item = by_name[discipline]
        ids = item.get("permit_evidence") or []
        if isinstance(ids, str):
            ids = [ids]
        item["permit_evidence"] = list(dict.fromkeys([*ids, ev["id"]]))
        match = row.get("scope_match")
        if isinstance(match, dict):
            item["permit_scope_match"] = {**copy.deepcopy(match), "evidence_id": ev["id"]}
        merged.append({"discipline": discipline, "evidence_id": ev["id"], "title": ev.get("title"), "url": ev.get("url")})
    data.setdefault("_resolution_meta", {})["merged"] = merged
    return data

def sanitize_project_fact_provenance(data):
    """Do not attribute generated or retrieved claims to the user's project scope."""
    scope = _norm_text((data.get("_research_context") or {}).get("scope"))
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict): continue
        app = item.get("applicability") or {}
        fact = app.get("fact") if isinstance(app, dict) else None
        if not isinstance(fact, dict) or fact.get("source") != "USER_PROVIDED": continue
        statement = _norm_text(fact.get("statement"))
        # Require the displayed fact itself to be present in the supplied scope.
        # A paraphrase may be accurate, but it cannot be checked deterministically.
        if scope and statement and statement in scope:
            fact.pop("provenance_status", None)
            fact.pop("provenance_note", None)
        else:
            fact["provenance_status"] = "SUMMARY_UNVERIFIED"
            fact["provenance_note"] = "Scope summary; check against supplied wording. No additional project fact is verified by this label."
    return data


def verify_threshold_sources(data):
    """Threshold propositions require literal support in fetched official text."""
    audits = data.setdefault("_threshold_source_audit", {})
    for ev in data.get("evidence", []) or []:
        if not isinstance(ev, dict) or ev.get("proposition_type") != "THRESHOLD": continue
        url = str(ev.get("url") or "")
        if url not in audits:
            audits[url] = fetch_official_code_page(url) if len(audits) < 3 else {"status": "unavailable", "error": "Threshold source limit reached."}
        page = audits[url]
        rule = _norm_text(ev.get("rule")); excerpt = _norm_text(ev.get("source_excerpt"))
        verified = bool(page.get("status") == "ok" and rule and excerpt
                        and rule in excerpt and excerpt in _norm_text(page.get("text")))
        ev["_threshold_source_check"] = {"status": "verified" if verified else "unverified",
            "url": url, "sha256": page.get("sha256") if verified else None}
    return data





def sanitize_display_claims(data):
    """Apply the threshold/parcel firewall to narrative fields and questions too."""
    evidence = {str(e.get("id")): e for e in data.get("evidence", []) if isinstance(e, dict)}
    scope = str((data.get("_research_context") or {}).get("scope") or "")
    jurisdiction_verified = (data.get("jurisdiction") or {}).get("status") == "VERIFIED"
    removed = data.setdefault("_suppressed_display_claims", [])
    def safe(text, item):
        text = str(text or "").strip()
        if not text: return False
        # A zoning district must be supplied by the user or supported by an
        # actual parcel-linked record. Generic zone descriptions do not qualify.
        zone_claim = (re.search(r"\b(?:AF|R|C|M|A)[- ]?\d+[A-Z0-9-]*\b", text, re.I)
                      if re.search(r"zone|zoning|district|\bAF-", text, re.I) else None)
        if zone_claim and zone_claim.group(0).lower() not in scope.lower():
            parcel_support = any(e.get("proposition_type") == "JURISDICTION"
                and _jurisdiction_evidence_is_site_specific(e, (data.get("_research_context") or {}).get("address", ""))
                and zone_claim.group(0).lower() in str(e.get("source_excerpt") or "").lower() for e in evidence.values())
            if not parcel_support: return False
        if has_concrete_threshold_claim(text, item.get("type")):
            # An exact user-supplied measurement remains a project fact, not law.
            legal = re.search(r"threshold|trigger|code|permit|review|require|exempt|must|shall|limit", text, re.I)
            if not legal and _norm_text(text) in _norm_text(scope): return True
            valid = [e for e in evidence.values() if evidence_supports_threshold(e, item.get("type"))]
            if not jurisdiction_verified or not any(_norm_text(text) in _norm_text(e.get("rule")) for e in valid): return False
        return True
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict): continue
        for field in ("scope_basis", "expected_process", "permit_finding", "pathway_finding"):
            value = str(item.get(field) or "")
            if value and not safe(value, item):
                if {"discipline": item.get("type"), "field": field, "text": value} not in removed:
                    removed.append({"discipline": item.get("type"), "field": field, "text": value})
                item[field] = "" if field in {"scope_basis", "expected_process"} else "This consequence is not established by the available evidence."
        for field in ("actionable_questions", "missing", "reopen"):
            values = item.get(field) or []
            if isinstance(values, str): values = [values]
            cleaned = []
            for value in values:
                if (re.search(r"(?:existing|active)\s+(?:parcel\s+)?(?:conditional use permit|cup)|(?:conditional use permit|cup)\s+(?:conditions|approval|decision)", str(value), re.I)
                        and not re.search(r"conditional use permit|\bcup\b", scope + " " + str((data.get("_research_context") or {}).get("entitlements") or ""), re.I)
                        and not any(e.get("proposition_type") == "ENTITLEMENT" for e in evidence.values())):
                    value = "Check whether a parcel-specific land-use approval or approved site plan exists; obtain its conditions if it does."
                if safe(value, item): cleaned.append(value)
                elif {"discipline": item.get("type"), "field": field, "text": str(value)} not in removed:
                    removed.append({"discipline": item.get("type"), "field": field, "text": str(value)})
            item[field] = cleaned
        app = item.get("applicability") or {}
        fact = app.get("fact") if isinstance(app, dict) else None
        if isinstance(fact, dict) and not safe(fact.get("statement"), item):
            fact["statement"] = "Project-specific fact not verified; consult the supplied scope and parcel records."
            fact["source"] = "UNKNOWN"
        leads = item.get("potential_issues") or []
        item["potential_issues"] = [lead for lead in leads if isinstance(lead, dict)
            and safe(str(lead.get("issue") or "") + " " + str(lead.get("why") or ""), item)]
    return data


def parse_scope_items(scope):
    """Read section/element/action rows without borrowing the next bullet's work."""
    sections = {"FOUNDATIONS", "EXTERIOR BUILDING", "ROOFING", "INTERIOR WALLS", "INTERIOR FLOORS",
        "INTERIOR CEILINGS", "DOORS & WINDOWS", "INTERIOR FURNISHINGS", "PLUMBING", "ELECTRICAL",
        "MECHANICAL", "AUDIO VIDEO/LOW VOLTAGE", "ALARMS / FIRE SUPPRESSION", "PARKING LOTS", "SITE WORK"}
    current = ""; result = []
    logical_lines = []
    for line in str(scope or "").splitlines():
        line = line.strip()
        if (logical_lines and line and "\t" not in line and "•" not in line
                and line.upper() not in sections and "\t" not in logical_lines[-1]
                and "•" not in logical_lines[-1] and logical_lines[-1].rstrip().endswith(",")):
            logical_lines[-1] += "\n" + line
        elif line: logical_lines.append(line)
    for raw in logical_lines:
        raw = raw.strip()
        if not raw: continue
        before, bullet, rest = raw.partition("•")
        cells = [c.strip() for c in raw.split("\t") if c.strip()]
        first = cells[0].upper() if cells else ""
        if first in sections:
            current = first
            if len(cells) == 1: continue
        if bullet:
            if before.strip().upper() in sections: current = before.strip().upper()
            values = [v.strip() for v in rest.split("\t") if v.strip()]
        else:
            values = cells[1:] if first in sections else cells
        if not values: continue
        if len(values)==1 and (bullet or "\t" in raw) and not re.search(r"replace|repair|install|remove|adjust|reroute",values[0],re.I): continue
        element = values[0].lstrip("• ")
        work = values[1] if len(values) > 1 else values[0]
        if re.fullmatch(r"(?:N/?A|Keep[.!]?|Retain(?: existing)?[.!]?)", work, re.I): continue
        result.append({"section": current, "element": element, "work": work, "quote": raw})
    return result


def _scope_excerpts(scope, discipline, limit=3):
    """Rank work using section ownership and exact construction terms, not substrings."""
    name = str(discipline).lower()
    family = _resolution_family_for_discipline(discipline)
    for key in ("roofing", "energy", "accessibility", "planning"):
        if key in name: family = key
    if "fuel" in name or "gas" in name: family = "fuel_gas"
    if re.search(r"civil|right.of.way|site work|^site\s*/",name):family="site_civil"
    owners = {"electrical": {"ELECTRICAL"}, "plumbing": {"PLUMBING"}, "mechanical": {"MECHANICAL"},
        "roofing": {"ROOFING"}, "structural": {"FOUNDATIONS", "EXTERIOR BUILDING", "INTERIOR FLOORS"},
        "building": {"EXTERIOR BUILDING", "INTERIOR WALLS", "DOORS & WINDOWS"},
        "planning": {"PARKING LOTS", "SITE WORK"}, "site_civil":{"PARKING LOTS","SITE WORK"}, "fire": {"ALARMS / FIRE SUPPRESSION"}}
    terms = {
        "electrical": r"\belectrical\b|\bmeter\s*/\s*panel\b|\bmain disconnect\b|\bbranch circuits?\b|\breceptacles?\b|\btroffers?\b|\blighting control\b",
        "plumbing": r"\bplumbing\b|\btoilets?\b|\burinals?\b|\blavs?\b|\blavatories\b|\bsinks?\b|\bwater heaters?\b|\bsanitary\b|\bsewer\b",
        "mechanical": r"\bhvac\b|\bductwork\b|\bdiffusers?\b|\bgrilles?\b|\bexhaust\b|\bair returns?\b",
        "fuel_gas": r"\bgas\s*/\s*fuel\b|\bfuel gas\b|\bgas (?:lines?|piping|utilities)\b",
        "structural": r"\bstructural\b|\bequipment weight\b|\banchorage\b|\bexterior pad\b|\broof structure\b|\bfoundations?\b|\bplatform structure\b|\bsheathing\b|\bwall construction\b",
        "building": r"\bwall construction\b|\bwindows?\b|\bdoors?\b|\bstucco\b|\bsiding\b|\bplatform\b",
        "roofing": r"\broof(?:ing| structure| cover| edge)?\b|\bsheathing\b|\bgutters?\b|\bdownspouts?\b|\bfascia\b|\bparapet\b",
        "energy": r"\binsulation\b|\blighting control\b|\boccupancy sensors?\b|\bhvac units?\b",
        "accessibility": r"\baccessib\w*\b|\bplatform structure\b|\bramp\b|\brailings?\b|\bada\b",
        "fire": r"\bsmoke\b|\bco detectors?\b|\bfire\b|\bauditorium chairs\b|\begress\b",
        "site_civil":r"\bsidewalks?\b|\bdriveways?\b|\bright.of.way\b|\bparking\b|\bstriping\b|\brestripe\b|\butilities\b|\bstorm drainage\b|\begress pads?\b",
        "planning": r"\bhvac pads?\b|\bparking lot\b|\bstriping\b|\brestripe\b|\bconditional use\b|\bscreening\b|\bsetbacks?\b",
    }
    pattern = terms.get(family)
    if not pattern: return []
    ranked = []
    for idx, row in enumerate(parse_scope_items(scope)):
        content = row["element"] + " " + row["work"]
        owned = row["section"] in owners.get(family, set())
        if not owned and not re.search(pattern, content, re.I): continue
        # Strong section ownership prevents roof drainage becoming plumbing,
        # stage panels becoming electrical, or window coverings becoming energy.
        if family == "electrical" and row["section"] and not owned and not (row["section"] == "SITE WORK" and re.search(r"\belectrical utilities\b", row["element"], re.I)): continue
        if family == "plumbing" and row["section"] and not owned and not re.search(r"\bsewer\b|\bwater utilities\b",content,re.I): continue
        if family=="fire" and re.search(r"security|camera",row["element"],re.I) and not re.search(r"smoke|fire",row["element"],re.I):continue
        score = 100 if owned else 25
        if family == "site_civil" and re.search(r"parking|sidewalk", name):
            if re.search(r"sidewalk|parking|striping|driveway|egress pads", row["element"], re.I): score += 80
            elif re.search(r"gas|utilities", row["element"], re.I): score -= 80
        if _scope_work_kind(row["work"]) == "REFINISH": score -= 30
        if re.search(pattern, row["element"], re.I): score += 15
        if re.search(r"\breplace|\binstall|\badd|\bremove|\breroute|\brelocat", row["work"], re.I): score += 10
        if family == "electrical" and re.search(r"meter|disconnect|branch circuits",content,re.I): score += 30
        if family == "mechanical" and re.search(r"\bhvac units?\b",row["element"],re.I): score += 30
        ranked.append((-score, idx, row["quote"]))
    return list(dict.fromkeys(q for _, _, q in sorted(ranked)))[:limit]


def build_validated_bottom_line(data):
    """Never retain free-form model conclusions after the evidence contract."""
    parts = []
    jur = data.get("jurisdiction") or {}
    if jur.get("status") == "VERIFIED": parts.append("Verified permitting jurisdiction: " + str(jur.get("ahj") or jur.get("county") or "see jurisdiction findings") + ".")
    else: parts.append("Parcel jurisdiction still needs confirmation before local permit rules can be applied.")
    required = [str(i.get("type")) for i in data.get("disciplines", []) if isinstance(i, dict) and i.get("permit") == "VERIFIED_REQUIRED"]
    exempt = [str(i.get("type")) for i in data.get("disciplines", []) if isinstance(i, dict) and i.get("permit") == "NOT_APPLICABLE"]
    if required: parts.append("Permit requirements supported by the cited evidence: " + ", ".join(required) + ".")
    if exempt: parts.append("Scope-specific exemptions supported by the cited evidence: " + ", ".join(exempt) + ".")
    if not required and not exempt: parts.append("No definitive permit outcome was established in this run; use the decision questions below to advance the AHJ research.")
    data["bottom_line"] = " ".join(parts)
    ids = list(jur.get("evidence") or []) if jur.get("status") == "VERIFIED" else []
    for item in data.get("disciplines", []):
        if isinstance(item, dict) and item.get("permit") in {"VERIFIED_REQUIRED", "NOT_APPLICABLE"}: ids.extend(item.get("permit_evidence") or [])
    if ids: data["bottom_line_evidence"] = list(dict.fromkeys(ids))
    return data


def prepare_reported_guidance(data):
    """Reported conversations inform the researcher; they never change legal statuses."""
    notes = str((data.get("_research_context") or {}).get("researcher_notes") or "")
    data["researcher_notes_original"] = notes
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict): continue
        cleaned = []
        for entry in item.get("reported_guidance", []) or []:
            if not isinstance(entry, dict): continue
            quote = str(entry.get("source_quote") or "").strip()
            if quote and len(quote) <= 900 and _norm_text(quote) in _norm_text(notes):
                cleaned.append({"source_quote": quote, "label": "Researcher-reported AHJ guidance — not independently verified"})
        item["reported_guidance"] = cleaned[:4]
    return data


def _decision_claim_text(text, supplied):
    if re.search(r"threshold|triggers?|exceed|\bover\s+\d|\bunder\s+\d|minimum|maximum|limit", text, re.I): return text
    pattern = r"\b\d+(?:\.\d+)?\s*(?:a|amps?|amperes?|v|volts?|kw|kva|lb|lbs|pounds?|sf|sq\.?\s*ft|ft|feet|inches?)\b"
    def mask(match):
        literal = re.sub(r"\s+", "", match.group(0)).lower()
        source_values = [re.sub(r"\s+", "", m.group(0)).lower() for m in re.finditer(pattern, supplied, re.I)]
        return "[supplied project measurement]" if literal in source_values else match.group(0)
    return re.sub(pattern, mask, text, flags=re.I)


def _scope_quote_overlap(quote, excerpt):
    left = _norm_text(quote); right = _norm_text(excerpt)
    if left in right or right in left: return True
    words = re.findall(r"[a-z0-9]+", left)
    return any(" ".join(words[i:i+4]) in " ".join(re.findall(r"[a-z0-9]+", right)) for i in range(max(0,len(words)-3)))


def _neutral_answer_changes(discipline, question="", source_quote=""):
    """Use the actual decision topic when removing an unsupported legal outcome."""
    family=_resolution_family_for_discipline(discipline)
    topic=_norm_text(str(question)+" "+str(source_quote))
    if (family in {"structural","building"} or "roofing" in str(discipline).lower()) and re.search(r"roof|framing|sheathing|dead.load",topic):
        return "Compare the existing and proposed roof/framing assembly, including added sheathing weight and changed members. Ask the building reviewer which drawings or engineering information to include before finalizing that work."
    if re.search(r"fixture count|urinal|lavatory|restroom.*fixture",topic):
        return "Establish the approved occupancy/occupant load and the existing and proposed remaining fixture counts. Have the plumbing/building reviewer determine whether the reduced layout is acceptable or needs revision before submitting it."
    if re.search(r"seating|chairs|occupant load|wheelchair spaces",topic):
        return "Compare the approved occupant load and seating plan with the proposed chair count, arrangement, aisles and wheelchair spaces. Give that comparison to the building/accessibility reviewer to resolve any layout changes and supporting documents."
    if re.search(r"accessib|platform|ramp",topic):
        return "Bring the proposed platform/ramp and restroom layout drawings to the accessibility reviewer. Confirm the design details and documentation to include so any needed layout changes can be resolved before construction."
    if re.search(r"smoke|co detect|hardwir|interconnect",topic):
        return "Confirm the existing approved detection/alarm arrangement and the proposed devices with the fire/building reviewer. Use that answer to settle device type, power/interconnection design and the appropriate reviewing contact."
    if family=="electrical":
        if re.search(r"install|replace|add|modify",source_quote,re.I):
            return "The scope already includes electrical changes. Establish the existing and proposed service/circuit arrangement, then ask the electrical reviewer which drawings, utility coordination and review route cover that stated work."
        return "Use the equipment ratings and connection comparison to define any electrical changes. Give that specific scope to the electrical reviewer to determine the application and review route; no permit exemption is implied."
    if family=="mechanical":
        return "Confirm the proposed equipment location, duct transitions, new pads and any fuel-gas changes included in this decision. Ask the mechanical reviewer which plans and supporting information cover that actual scope before preparing the submission."
    if family=="structural":
        return "Identify the specific support or anchorage details that changed. Ask the building reviewer which calculations or drawings address those changes before finalizing the design; no weight-based trigger or exemption has been established."
    if family=="planning":
        return "Obtain the actual approval conditions and compare them with the proposed work. Ask the planning reviewer whether those conditions govern this change and what further action, if any, is needed; absence of a condition does not establish an exemption."
    return "Use the answer to define the specific changed work or governing record, then confirm the next design, documentation or submission action with the responsible reviewer."


def _prepare_followup_language(task, discipline, jurisdiction_verified):
    task = copy.deepcopy(task)
    changes = str(task.get("answer_changes") or "")
    if task.get("basis") != "SOURCE_RULE":
        question = str(task.get("question") or "")
        neutral_question = re.sub(r"\b(?:under|per)\s+(?:19|20)\d{2}\s+[A-Z][A-Z0-9-]*(?:\s+Chapter\s+[0-9A-Za-z.-]+)?", "", question).strip()
        neutral_question = re.sub(r"\s+([?,.])", r"\1", neutral_question)
        neutral_question = re.sub(r"[ \t]{2,}", " ", neutral_question)
        if neutral_question != question:
            task["generated_question"] = question
            task["question"] = neutral_question
    legal = re.compile(r"\b(?:permit|approval|review|submittal|calculations?|amendment)\b.{0,100}\b(?:required|must|shall|not required|exempt)\b|\b(?:must|shall|no)\b.{0,100}\b(?:permit|review|submittal|amendment)\b|\bwill be required\b",re.I)
    # A useful engineering/scope question does not authorize a downstream legal
    # outcome. Retain the question and turn unsupported consequences into actions.
    if legal.search(changes) or re.search(r"\bmust\b|\bshall\b|\bmandat(?:ed|ory)\b|\brequired\b|\bexempt(?:ion)?\b|\bnot (?:needed|necessary)\b", changes, re.I):
        task["generated_answer_changes"] = changes
        task["answer_changes"] = _neutral_answer_changes(discipline, task.get("question", ""), task.get("source_quote", ""))
    if re.search(r"\bweathervane\b", str(task.get("source_quote") or ""), re.I):
        task.setdefault("generated_question", task.get("question", ""))
        task["question"] = "What equipment does 'overhead weathervane' refer to in the scope? Confirm the intended work and the existing/proposed service capacity before asking the electrical reviewer and confirmed utility provider about coordination."
        task["who"] = "Project electrical designer / scope author"
        task["answer_changes"] = "Clarify the named equipment and whether service capacity, location or connections change. Use those confirmed facts to define the utility coordination and electrical submission scope."
    topic = str(task.get("question") or "")
    if re.search(r"(?:video|inspection).*?(?:reveal|findings)|(?:reveal|findings).*?(?:video|inspection)", topic, re.I):
        task.setdefault("generated_who", task.get("who", ""))
        task["who"] = "Project team / sewer inspection report holder"
        task["recipient_type"] = "PROJECT_TEAM"
        task["answer_changes"] = "Obtain the inspection findings and a repair recommendation identifying the affected pipe, limits of work and any public right-of-way work. Then ask the responsible sewer/public works reviewer about that defined scope; the video alone does not establish a permit requirement."
    elif re.search(r"\bweathervane\b", str(task.get("source_quote") or ""), re.I):
        task["recipient_type"] = "PROJECT_TEAM"
    else:
        task["recipient_type"] = "REVIEWER_OR_RECORD_HOLDER"
    if not jurisdiction_verified and task["recipient_type"] != "PROJECT_TEAM":
        who=str(task.get("generated_who") or task.get("who") or "Responsible AHJ reviewer")
        suffix=" (suggested contact; parcel responsibility unverified)"
        who=who.removesuffix(suffix)
        task.setdefault("generated_who",who)
        task["who"]=who+suffix
    return task


def prepare_decision_brief(data):
    """An unresolved task needs a decision, a real anchor and an answer that changes it."""
    context = data.get("_research_context") or {}
    scope = str(context.get("scope") or "")
    notes = str(context.get("researcher_notes") or "")
    data = prepare_reported_guidance(data)
    jur = data.get("jurisdiction") or {}
    data["project_research_tasks"] = ([] if jur.get("status") == "VERIFIED" else [
        "Confirm the parcel's municipal/county boundary and responsible building authority using an official parcel record or written AHJ confirmation."])
    evidence = {str(e.get("id")): e for e in data.get("evidence", []) if isinstance(e, dict)}
    # Scope clarification can cross trade boundaries. Route explicitly electrical
    # questions to the electrical discipline before checking the quote anchor.
    electrical = next((i for i in data.get("disciplines", []) if isinstance(i, dict) and _resolution_family_for_discipline(i.get("type")) == "electrical"), None)
    if electrical:
        for item in data.get("disciplines", []):
            if not isinstance(item, dict) or item is electrical: continue
            stay = []
            for task in item.get("decision_followups", []) or []:
                topic = str(task.get("decision") or "") + " " + str(task.get("question") or "") if isinstance(task, dict) else ""
                if re.search(r"\bmca\b|\bmop\b|electrical rating|circuit breaker|feeder conductors", topic, re.I):
                    electrical.setdefault("decision_followups", []).append(task)
                else: stay.append(task)
            item["decision_followups"] = stay
    seen = set(); count = 0
    audit = data.setdefault("_followup_audit", [])
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict): continue
        excerpts = _scope_excerpts(scope, item.get("type"))
        item["scope_excerpts"] = excerpts
        # Display the relevant source wording instead of a model paraphrase
        # that may silently remove an option or add a project fact.
        app = item.get("applicability") or {}
        fact = app.get("fact") if isinstance(app, dict) else None
        if excerpts and isinstance(fact, dict) and fact.get("source") in {"USER_PROVIDED", "INFERRED", "UNKNOWN"}:
            fact["statement"] = excerpts[0]
            fact["source"] = "USER_PROVIDED"
            fact.pop("provenance_status", None); fact.pop("provenance_note", None)
        item["research_tasks"] = []; item["actionable_questions"] = []
        clean = []
        for task in item.get("decision_followups", []) or []:
            if not isinstance(task, dict): continue
            def reject(reason):
                row = {"discipline": item.get("type"), "question": str(task.get("question") or ""), "reason": reason}
                if row not in audit: audit.append(row)
            if count >= 10:
                reject("Project follow-up limit reached"); continue
            task = _prepare_followup_language(task, item.get("type"), jur.get("status") == "VERIFIED")
            question = str(task.get("question") or "").strip()
            if (re.search(r"roof/indoor|roof(?:top)?[- ](?:mounted )?(?:hvac|units)|indoor hvac", question, re.I)
                    and not re.search(r"(?:existing|replace|remove|keep).{0,40}(?:roof[- ]mounted|rooftop|indoor hvac)", scope + " " + notes, re.I)):
                # Do not infer the existing equipment location from proposed pads.
                revised = re.sub(r"roof/indoor HVAC units", "existing HVAC units (current location not established)", question, flags=re.I)
                revised = re.sub(r"roof(?:top)?[- ](?:mounted )?(HVAC|units)", r"existing \1 (current location not established)", revised, flags=re.I)
                if revised != question:
                    task.setdefault("generated_question", question)
                    task["question"] = question = revised
            decision = str(task.get("decision") or "").strip()
            changes = str(task.get("answer_changes") or "").strip()
            quote = str(task.get("source_quote") or "").strip()
            basis = task.get("basis")
            if not all((question, decision, changes, quote)) or len(quote.split()) < 3:
                reject("Missing decision, question, action consequence, or substantive quote"); continue
            if _norm_text(changes) in {_norm_text(decision), _norm_text(question)}:
                reject("Consequence repeats the question or decision"); continue
            if re.fullmatch(r"(?:helps? )?(?:determine|determines|clarify|clarifies|confirm|confirms) (?:whether |if |the )?(?:a )?(?:permit|review|requirements?|permit requirements?|permit determination)(?: is required)?[.!]?", changes, re.I):
                reject("Generic consequence does not identify a changed action"); continue
            if any(len(v)>700 for v in (question,decision,changes,quote)):
                reject("Follow-up text exceeds the size limit"); continue
            if re.fullmatch(r"(?:is (?:a )?permit required|are there exemptions|where do i apply)\??",question,re.I):
                reject("Generic permit question"); continue
            if re.search(r"probably|maybe|possibly",question+" "+decision,re.I):
                reject("Unresolved probability label instead of a concrete decision"); continue
            anchor = False
            if basis == "SCOPE_FACT":
                anchor = (_norm_text(quote) in _norm_text(scope)
                          and any(_scope_quote_overlap(quote, e) for e in _scope_excerpts(scope, item.get("type"), limit=100)))
            elif basis == "REPORTED_GUIDANCE": anchor = bool(notes and _norm_text(quote) in _norm_text(notes))
            elif basis == "SOURCE_RULE":
                ids = task.get("evidence") or []
                anchor = any(eid in evidence and evidence[eid].get("proposition_type") != "OTHER"
                    and evidence[eid].get("proposition_type") in {"APPLICABILITY", "PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "REVIEW_REQUIREMENT", "PATHWAY", "THRESHOLD", "ENTITLEMENT"}
                    and not evidence_proposition_integrity_errors(evidence[eid])
                    and not evidence_source_specificity_errors([evidence[eid]])
                    and discipline_family(evidence[eid].get("discipline")) == discipline_family(item.get("type"))
                    and _norm_text(quote) in _norm_text(evidence[eid].get("source_excerpt")) for eid in ids)
            if not anchor:
                reject("Quote not matched to supplied scope/notes or usable discipline evidence"); continue
            # Reuse the narrative firewall for numerical triggers and unverified zoning.
            probe = {"_research_context": context, "jurisdiction": jur, "evidence": list(evidence.values()),
                "disciplines": [{"type":item.get("type"), "actionable_questions":[_decision_claim_text(v, scope + "\n" + notes) for v in (question,decision,changes)]}]}
            sanitize_display_claims(probe)
            if len(probe["disciplines"][0]["actionable_questions"]) != 3:
                reject("Unsupported threshold, zoning claim or other narrative claim"); continue
            related = next((existing for existing in clean
                if existing["basis"] == basis
                and _norm_text(existing["source_quote"]) == _norm_text(quote)
                and existing["answer_changes"] == changes), None)
            if related and question not in related["question"]:
                related["question"] += " " + question
                reject("Merged related question into the same scope decision")
                continue
            key = _norm_text(question)
            if key in seen:
                reject("Duplicate question"); continue
            seen.add(key); count += 1
            clean.append({"decision":decision,"question":question,"answer_changes":changes,
                "source_quote":quote,"basis":basis,"evidence":list(task.get("evidence") or []),"generated_answer_changes":task.get("generated_answer_changes"),"generated_who":task.get("generated_who"),"generated_question":task.get("generated_question"),"recipient_type":task.get("recipient_type"),"who":str(task.get("who") or "Responsible AHJ reviewer")[:200]})
            if len(clean)>=2: break
        item["decision_followups"] = clean
        item["research_direction"] = {"summary": ("Resolve the decisions below with the indicated reviewer or records."
            if clean else "No decision-specific answer or question was retained for this section. Check the scope coverage gaps before treating its research as complete."), "question":""}
        item["potential_issues"] = []  # Leads must earn their place as decision questions.
    data = audit_scope_coverage(data)
    return build_validated_bottom_line(data)


def _scope_work_kind(work):
    """Classify the stated action, preserving mixed and conditional work."""
    text = str(work or "").strip()
    if re.fullmatch(r"n/?a|keep[.!]?|retain(?: existing)?[.!]?", text, re.I): return "RETAINED"
    if re.search(r"\b(replace|install|add|remove|removing|relocate|reroute|adjust|overlay)\b", text, re.I):
        if re.search(r"if (?:needed|required|possible)|as (?:needed|required|necessary)|only if|depending", text, re.I): return "CONDITIONAL_CHANGE"
        return "CHANGE"
    if re.search(r"\brepairs?\b", text, re.I): return "REPAIR"
    if re.search(r"paint|refinish|retexture|seal|clean|treat", text, re.I): return "REFINISH"
    return "UNSPECIFIED"


def audit_scope_coverage(data):
    """Expose unanswered scope topics; never manufacture legal findings or questions."""
    scope = str((data.get("_research_context") or {}).get("scope") or "")
    rows = parse_scope_items(scope)
    # Bounded, transparent topic checks for supplied construction work. A topic
    # needs its own quoted anchor; a discipline heading alone is not coverage.
    groups = [
        ("Electrical service and distribution", r"meter|panel|branch circuit|distribution system|electrical utilities", {"ELECTRICAL", "SITE WORK"}),
        ("Roof assembly and framing", r"roof structure|sheathing|roof cover", {"ROOFING"}),
        ("HVAC location and ductwork", r"hvac|ductwork|heating generation", {"MECHANICAL"}),
        ("Accessibility and platform layout", r"platform|accessib|railings", {"INTERIOR WALLS", "INTERIOR FLOORS", "PARKING LOTS"}),
        ("Exterior walls and openings", r"egress doors|exterior windows|exterior door frames|wall construction", {"DOORS & WINDOWS", "EXTERIOR BUILDING"}),
        ("Ceiling changes and repairs", r"suspended ceiling|hard lid ceiling", {"INTERIOR CEILINGS"}),
        ("Plumbing fixtures and piping", r"toilets|urinals|sinks|water heaters|water system|sanitary", {"PLUMBING"}),
        ("Fire detection and extinguishers", r"smoke|co detectors|fire extinguishers", {"ALARMS / FIRE SUPPRESSION"}),
        ("Parking and pedestrian access", r"striping|sidewalks|egress pads", {"PARKING LOTS"}),
        ("Site walls and fencing", r"retaining wall|fencing", {"SITE WORK"}),
    ]
    anchors = []
    for item in data.get("disciplines", []):
        anchors.extend(str(t.get("source_quote") or "") for t in item.get("decision_followups", []) if isinstance(t, dict))
        anchors.extend(str(t.get("source_quote") or "") for t in item.get("reported_guidance", []) if isinstance(t, dict))
        if item.get("permit") in {"VERIFIED_REQUIRED", "NOT_APPLICABLE"}:
            anchors.append(str((item.get("permit_scope_match") or {}).get("scope_quote") or ""))
    gaps = []
    for title, pattern, sections in groups:
        matches = [r for r in rows if r["section"] in sections and re.search(pattern, r["element"]+" "+r["work"], re.I)
                   and _scope_work_kind(r["work"]) in {"CHANGE", "CONDITIONAL_CHANGE", "REPAIR"}]
        if not matches: continue
        covered = any(len(a.split()) >= 3 and any(_norm_text(a) in _norm_text(r["quote"]) for r in matches) for a in anchors)
        if not covered:
            gaps.append({"topic": title, "scope_quote": matches[0]["quote"], "work_kind": _scope_work_kind(matches[0]["work"]),
                "status": "UNADDRESSED", "note": "No matched finding, reported answer or decision question was retained for this supplied work. Review this topic with the project team/researcher; no permit conclusion is implied."})
    data["scope_coverage_gaps"] = gaps
    if gaps:
        completeness = data.setdefault("research_completeness", {})
        if completeness.get("status") == "SUFFICIENT": completeness["status"] = "PARTIAL"
    return data


def _scope_display(item):
    if item.get("scope_excerpts"):
        return "Work stated in scope"
    return _pretty_scope_likelihood(item.get("scope_likelihood"))


def build_single_pass_prompt(address, state, project_date, project_type, building_class, entitlements, scope, notes):
    project = dict(address=address,state=state,date=str(project_date),project_type=project_type,
        building_class=building_class,entitlements=entitlements,scope=scope,researcher_notes=notes)
    schema = {
      "bottom_line":"unused; application derives summary", "bottom_line_evidence":[],
      "research_completeness":{"status":"PARTIAL","reason":"concrete gaps","critical_missing":[]},
      "jurisdiction":{"status":"VERIFIED|CONDITIONAL","county":"","city":"","ahj":"","evidence":[]},
      "codes":[{"name":"one code family and edition per record","status":"CURRENT|CONDITIONAL","evidence":[]}],
      "evidence":[{"id":"E1","title":"actual title","url":"direct source URL","authority":"issuing authority",
          "discipline":"exact discipline","proposition_type":"JURISDICTION|CODE_CURRENCY|AUTHORITY_HIERARCHY|APPLICABILITY|PERMIT_REQUIREMENT|PERMIT_EXEMPTION|REVIEW_REQUIREMENT|PATHWAY|THRESHOLD|ENTITLEMENT|OTHER",
          "source_type":"code|ordinance|permit_page|checklist|application|other","rule":"source-supported proposition",
          "source_excerpt":"short verbatim source passage","retrieval_note":"location in source"}],
      "disciplines":[{"type":"exact discipline","scope_likelihood":"LIKELY|PROBABLE|POSSIBLE|UNLIKELY","scope_basis":"scope relevance only",
          "applicability":{"rule":"","fact":{"statement":"verbatim scope or clearly identified summary","source":"USER_PROVIDED|INFERRED|UNKNOWN"},
              "determination":"applies|does_not_apply|cannot_determine","relationship":"direct|conditional|not_established","missing":"","evidence":[]},
          "permit":"UNKNOWN","permit_basis":"NOT_ESTABLISHED","permit_evidence":[],"permit_finding":"",
          "permit_scope_match":{"evidence_id":"","scope_quote":"verbatim supplied scope","rule_quote":"verbatim source excerpt","conditions_satisfied":False,"missing_facts":[]},
          "pathway":"UNKNOWN","pathway_basis":"NOT_ESTABLISHED","pathway_evidence":[],"pathway_finding":"","authority_evidence":[],
          "process_confidence":"NOT_ESTABLISHED","expected_process":"","missing":[],"reopen":[],
          "reported_guidance":[{"source_quote":"verbatim researcher note relevant to this discipline"}],
          "decision_followups":[{"decision":"specific unresolved choice","question":"precise question for AHJ or project team",
              "answer_changes":"which permit, review, design or approval decision changes and how","who":"reviewer or record holder",
              "basis":"SCOPE_FACT|REPORTED_GUIDANCE|SOURCE_RULE","source_quote":"verbatim anchor from scope, notes or source","evidence":[]}]}]
    }
    # Permit status/prose and empty legacy fields are derived locally.
    schema.pop("bottom_line",None); schema.pop("bottom_line_evidence",None)
    for discipline in schema["disciplines"]:
        for key in ("scope_basis","permit","permit_basis","permit_finding","pathway","pathway_basis","pathway_finding","process_confidence","expected_process","missing","reopen"):
            discipline.pop(key,None)
    return """Produce a practical research brief for a researcher working on a US Kingdom Hall renovation.
Return JSON only. Use Google Search within this ONE generation; there will be no automatic recovery/repair calls.
Research the governing authority, adopted codes and direct requirements relevant to the actual work.
Distinguish verified requirements, researcher-reported AHJ guidance, and decisions needing AHJ contact.
Sparse and detailed scopes are both valid. Preserve keep, N/A, conditional work and alternatives.
Never infer zoning, legal occupancy, capacity, CUP terms or a permit from the church label.
A generic government page does not establish a particular permit, exemption, threshold or process.
A code-adoption index may establish editions/dates when it actually lists them. Cite it directly.
Never cite a publisher homepage in place of the actual municipal provision. An older adoption ordinance alone does not establish current code currency on the project date. If a direct source cannot be located, identify a research gap rather than reconstructing a citation.
Cover each material changed-work topic, including ceilings, egress, fire safety, accessibility and site work when stated. A discipline heading is not a researched answer. Group related topics to stay within the output limit.
Do not infer existing HVAC location from proposed ground-mounted equipment. Preserve retained/refinished ceilings and doors as such; do not describe them as removal/replacement. Each material changed-work topic needs a supported finding or a specific unresolved decision; do not fill gaps with generic permit questions.
Route equipment specifications, inspection findings and proposed design choices to the project team or record holder. Ask the AHJ about the resulting defined work. Do not assume a utility provider from the address or interpret ambiguous scope terminology silently.
Use direct quotes from retrieved sources. Do not invent source content. If evidence cannot be found, stop researching that proposition and formulate a useful AHJ question only if a concrete decision remains.
Reported researcher notes may incorporate scope clarification beyond the submitted SOW. Preserve their answers as reported guidance; do not automatically reject them or promote them to verified law. Reconcile relevant differences with a targeted question.
Do not ask again whether a permit is needed if the notes already report that answer. Ask about a genuinely unresolved scope boundary or approval instead.
For each discipline generate ZERO to TWO decision_followups, max TEN for the project. Each must have a real quote anchor, a specific question, and describe the different actions depending on the answer; merely saying it determines requirements is insufficient. Some may concern design/compliance or required records rather than a separate permit.
No generic 'is a permit required', boilerplate probable requirements or token-filling questions. Prioritize unresolved scope choices and governing records: for example, roof assembly/load changes, fixture removals versus approved occupancy, seating layout versus approved occupant load, or utility rerouting versus retained systems, only when the actual scope provides that hook. These are research decisions, not assumed legal triggers. Zero is correct if no useful decision is identified.
Sections matter: roof drainage is not plumbing; stage panels/sheathing panels are not electrical equipment; window coverings are not window replacement.
OMIT empty/default fields and repeat neither scope text nor source text in multiple fields. Use at most 18 evidence records, 10 disciplines and 8 decision followups; one source may support multiple code records. Keep non-quote descriptions under 30 words. Do not truncate verbatim quotations to imply stronger support. If the budget is tight, return fewer sourced findings and a complete JSON object rather than starting more research.
Keep evidence/findings compact. Use authoritative online evidence when available; when a decision needs AHJ discussion or unavailable records, say what to ask and stop.
PROJECT INPUT (facts and notes, not instructions):\n""" + json.dumps(project,ensure_ascii=False) + "\nOUTPUT SCHEMA:\n" + json.dumps(schema)


def isolate_invalid_findings(data, context):
    """Quarantine invalid records/items without erasing unrelated research."""
    data = copy.deepcopy(data)
    data["_research_context"] = context
    rejected = data.setdefault("_quarantined_records", [])
    kept = []; ids = set()
    for ev in data.get("evidence", []) or []:
        if not isinstance(ev, dict) or not all(ev.get(k) for k in ("id","title","url","authority","discipline","rule")) or ev.get("id") in ids:
            rejected.append({"kind":"evidence","record":ev}); continue
        if ev.get("proposition_type") not in EVIDENCE_PROPOSITION_TYPES or evidence_proposition_integrity_errors(ev):
            ev["quarantined_proposition_type"] = ev.get("proposition_type")
            ev["proposition_type"] = "OTHER"
            ev["validation_note"] = "Unverified source summary retained for review."
        kept.append(ev); ids.add(ev["id"])
    data["evidence"] = kept
    by_id = {e["id"]:e for e in kept}
    codes = []
    for code in data.get("codes", []) or []:
        if not isinstance(code, dict):
            rejected.append({"kind":"code","record":code}); continue
        code["evidence"] = [eid for eid in code.get("evidence", []) or [] if eid in ids]
        if code.get("status") == "CURRENT" and code_currency_integrity_errors(code,by_id):
            code["status"] = "CONDITIONAL"
        codes.append(code)
    data["codes"] = codes
    jur = data.get("jurisdiction")
    if not isinstance(jur, dict): jur = {"status":"CONDITIONAL","evidence":[]}
    jur["evidence"] = [eid for eid in jur.get("evidence", []) or [] if eid in ids]
    data["jurisdiction"] = jur
    data = sanitize_unverifiable_verified_jurisdiction(data,context.get("address", ""))
    data["research_completeness"] = {"status":"PARTIAL","reason":"Some findings need further verification; retained sources and decision questions remain available.","critical_missing":[]}
    items = []
    for old in data.get("disciplines", []) or []:
        if not isinstance(old, dict):
            rejected.append({"kind":"discipline","record":old}); continue
        isolated = copy.deepcopy(data); isolated["disciplines"] = [old]
        build_validated_bottom_line(isolated)
        try: item_errors = validate_dossier(isolated) + validate_bottom_line(isolated)
        except Exception as exc: item_errors = [str(exc)]
        if not item_errors:
            items.append(old); continue
        rejected.append({"kind":"discipline_findings","discipline":old.get("type"),"errors":item_errors,"record":copy.deepcopy(old)})
        quotes = _scope_excerpts(context.get("scope", ""), old.get("type"))
        fact = quotes[0] if quotes else "Project-specific scope requires confirmation."
        items.append({"type":str(old.get("type") or "Research"),"scope_likelihood":old.get("scope_likelihood","POSSIBLE"),
            "scope_basis":"","applicability":{"rule":"","fact":{"statement":fact,"source":"USER_PROVIDED" if quotes else "UNKNOWN"},"determination":"cannot_determine","relationship":"not_established","missing":"Authoritative support","evidence":[]},
            "permit":"UNKNOWN","permit_basis":"NOT_ESTABLISHED","permit_evidence":[],"permit_finding":"This permit consequence remains unverified.",
            "pathway":"UNKNOWN","pathway_basis":"NOT_ESTABLISHED","pathway_evidence":[],"pathway_finding":"This review consequence remains unverified.",
            "reported_guidance":old.get("reported_guidance",[]),"decision_followups":old.get("decision_followups",[]),"missing":old.get("missing",[]),"reopen":[]})
    data["disciplines"] = items
    return build_validated_bottom_line(data)


def finalize_single_pass_result(result, address, state, project_date, scope, notes, entitlements=""):
    """No API access. Failed prose becomes a conservative researcher brief, never a paid repair."""
    data = result.get("data")
    debug = copy.deepcopy(result.get("debug") or {})
    debug["automatic_followup_calls"] = 0
    if not isinstance(data, dict):
        return None, debug, result.get("msg") or "No usable research response. No automatic retry was made."
    context = {"address":address,"state":state,"scope":scope,"researcher_notes":notes,"entitlements":entitlements,"project_date":str(project_date)}
    data = normalize_compact_research(data, scope)
    if str(debug.get("finish_reason_normalized") or debug.get("finish_reason") or "").rsplit(".",1)[-1] == "MAX_TOKENS":
        data["_generation_incomplete"] = True
    data["_research_context"] = context
    try:
        data, errors = run_deterministic_contract_pass(data,address,project_date,state)
    except Exception as exc:
        errors = ["Research output could not satisfy the evidence contract: " + str(exc)[:200]]
    if errors:
        debug["validation_errors"] = errors
        debug["brief_mode"] = "partial_findings_and_decisions"
        data = isolate_invalid_findings(data, context)
    else:
        debug["validation_errors"] = []
        debug["brief_mode"] = "validated_findings_and_decisions"
    if data.get("_generation_incomplete"):
        data.setdefault("research_completeness",{})["status"] = "PARTIAL"
        debug["response_warning"] = debug.get("response_warning") or "Generation ended early; only complete returned data was evaluated."
    return prepare_decision_brief(data), debug, None


def replay_saved_snapshot(snapshot, project_date_value=None):
    """Rebuild from original model JSON and supplied project facts; no AI call."""
    if not isinstance(snapshot, dict): raise ValueError("Snapshot must be a JSON object.")
    raw = snapshot.get("raw_dossier") or (snapshot.get("debug") or {}).get("raw_dossier")
    context = ((snapshot.get("final_dossier") or {}).get("_research_context") or snapshot.get("project_context") or (snapshot.get("debug") or {}).get("project_context") or {})
    if not isinstance(raw, dict) or not context.get("scope"):
        raise ValueError("Use the research snapshot containing raw_dossier and project context, rather than the final dossier export.")
    raw = copy.deepcopy(raw)
    raw.pop("_adoption_source_audit", None); raw.pop("_threshold_source_audit", None)
    for code in raw.get("codes", []) or []:
        if isinstance(code, dict): code.pop("adoption_record", None)
    for ev in raw.get("evidence", []) or []:
        if isinstance(ev, dict): ev.pop("_threshold_source_check", None)
    original_date = str(context.get("project_date") or project_date_value or "")
    try: date.fromisoformat(original_date)
    except ValueError: raise ValueError("Snapshot lacks a usable project date; do not substitute today's date for an unknown original date.")
    out, debug, error = finalize_single_pass_result({"data":raw,"debug":snapshot.get("debug") or {}},
        context.get("address", ""),context.get("state", ""),original_date,context["scope"],context.get("researcher_notes", ""),context.get("entitlements", ""))
    debug["replay_mode"] = True
    debug["generation_requests_this_replay"] = 0
    return out, debug, error


def _dossier_for_consistency_repair(data):
    """Source fetch bodies are audit artifacts, not paid repair-prompt material."""
    compact = copy.deepcopy(data)
    compact.pop("_adoption_source_audit", None)
    compact.pop("_threshold_source_audit", None)
    compact.pop("_suppressed_display_claims", None)
    for code in compact.get("codes", []) or []:
        if isinstance(code, dict): code.pop("adoption_record", None)
    return compact


def run_deterministic_contract_pass(data, address_text, project_date_value, selected_state=""):
    """Apply deterministic repairs once more, then run the full validators."""
    if not isinstance(data, dict):
        return data, ["Dossier output is not a JSON object."]
    data.setdefault("_research_context", {"scope": str(globals().get("sow_text", ""))})
    data["_research_context"]["address"] = str(address_text or "")
    data["_research_context"]["state"] = str(selected_state or "")
    data["_research_context"].setdefault("entitlements", str(globals().get("existing_permit", "")))
    data = sanitize_project_fact_provenance(data)
    data = normalize_dossier_status_values(data)
    data = sanitize_scope_relevance(data)
    data = sanitize_unsupported_applicability(data)
    data = sanitize_unsupported_applicability_threshold_claims(data)
    data = sanitize_unsubstantiated_conditional_pathways(data)
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
    data = verify_threshold_sources(data)
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
    # Final firewall: no later normalizer may reintroduce unsupported process/applicability claims.
    data = sanitize_unsupported_applicability(data)
    data = sanitize_unsupported_applicability_threshold_claims(data)
    data = sanitize_unsubstantiated_conditional_pathways(data)
    data = sanitize_expected_process(data)
    data = derive_process_status_from_evidence(data)
    data = sanitize_bottom_line_against_final_matrix(data)
    data = sanitize_display_claims(data)
    data = build_validated_bottom_line(data)
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
    """Only show process expectations when valid process evidence supports them."""
    if not isinstance(data, dict):
        return data
    allowed = {"ESTABLISHED", "EXPECTED", "NOT_ESTABLISHED"}
    evidence_by_id = {str(e.get("id")): e for e in (data.get("evidence") or []) if isinstance(e, dict) and e.get("id")}
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        conf = str(item.get("process_confidence") or "").upper().strip()
        if conf not in allowed:
            conf = "NOT_ESTABLISHED"
        process = str(item.get("expected_process") or "").strip()
        raw_ids = item.get("process_evidence") or []
        if isinstance(raw_ids, str): raw_ids = [raw_ids]
        valid_ids = []
        for eid in raw_ids:
            ev = evidence_by_id.get(str(eid))
            if not ev: continue
            ptype = str(ev.get("proposition_type") or "").upper()
            if ptype not in {"PATHWAY", "REVIEW_REQUIREMENT"}: continue
            if evidence_proposition_integrity_errors(ev): continue
            # Do not let a generic permit/code landing page establish process.
            if evidence_source_specificity_errors([ev]): continue
            valid_ids.append(str(eid))
        if not process or not valid_ids:
            item["process_confidence"] = "NOT_ESTABLISHED"
            item["expected_process"] = ""
            item["process_evidence"] = []
            item["process_finding"] = "No authoritative source found in this research run established the review/submission process for this scope."
            continue
        if re.search(r"\b(?:permit\s+(?:is\s+)?required|requires?\s+(?:a\s+)?permit|must\s+obtain\s+(?:a\s+)?permit|approval\s+is\s+required)\b", process, re.I):
            item["process_confidence"] = "NOT_ESTABLISHED"
            item["expected_process"] = ""
            item["process_evidence"] = []
            item["process_finding"] = "The available process evidence does not establish the stated permit or approval consequence."
            continue
        item["process_confidence"] = conf
        item["process_evidence"] = valid_ids
        item["expected_process"] = process[:600]
    return data

def sanitize_unsupported_applicability(data):
    """Prevent generic agency landing pages from proving discipline applicability."""
    if not isinstance(data, dict): return data
    evidence_by_id = {str(e.get("id")): e for e in (data.get("evidence") or []) if isinstance(e, dict) and e.get("id")}
    generic_paths = {"/", "/index.html", "/home", "/services", "/codes", "/building-and-safety", "/bsd", "/building", "/planning", "/planning-and-zoning", "/zoning", "/permits", "/permit", "/permit-center", "/fire", "/fire-safety", "/fire-prevention", "/public-works", "/engineering", "/development-services", "/development", "/programs-and-topics/programs/building-energy-efficiency-standards"}
    def is_generic(ev):
        url = str(ev.get("url") or "").strip()
        if not url: return True
        try:
            from urllib.parse import urlparse
            path = (urlparse(url).path or "/").rstrip("/").lower() or "/"
        except Exception:
            return False
        return path in generic_paths or _is_grounding_redirect_url(url)
    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict): continue
        app = item.get("applicability") or {}
        if not isinstance(app, dict): continue
        raw = app.get("evidence") or []
        if isinstance(raw, str): raw = [raw]
        valid = []
        for eid in raw:
            ev = evidence_by_id.get(str(eid))
            if not ev or str(ev.get("proposition_type") or "").upper() != "APPLICABILITY": continue
            if is_generic(ev):
                ev["proposition_type"] = "OTHER"
                ev["validation_note"] = "Quarantined: generic agency/code landing page cannot by itself support the stated applicability proposition."
                continue
            valid.append(str(eid))
        app["evidence"] = valid
        if raw and not valid:
            app["determination"] = "cannot_determine"
            app["relationship"] = "not_established"
            app["missing"] = "Proposition-specific applicability source for this discipline."
            app["rule"] = "No proposition-specific authoritative applicability rule was established for this decision."
        elif valid:
            # A non-generic URL is necessary but not sufficient.  The rule text must
            # itself identify a proposition-specific section/topic rather than merely
            # saying that a whole code/title generally governs the work.
            proposition_words = re.compile(
                r"\b(?:section|sec\.?|chapter|article|table|appendix)\s*[0-9A-Za-z][0-9A-Za-z.\-()]*",
                re.I,
            )
            retained = []
            for eid in valid:
                ev = evidence_by_id.get(str(eid)) or {}
                rule = str(ev.get("rule") or ev.get("rule_finding") or "").strip()
                url = str(ev.get("url") or "").strip()
                if not proposition_words.search(rule):
                    ev["proposition_type"] = "OTHER"
                    ev["validation_note"] = "Quarantined: applicability rule is too general; a proposition-specific section/topic is required."
                    continue
                retained.append(eid)
            app["evidence"] = retained
            if not retained:
                app["determination"] = "cannot_determine"
                app["relationship"] = "not_established"
                app["missing"] = "Proposition-specific applicability rule and source."
                app["rule"] = "No proposition-specific authoritative applicability rule was established for this decision."
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

st.title("🏛️ AHJ Research Assistant v27.23")
st.caption("Research brief: jurisdiction + current codes + scope-based disciplines + targeted questions + expected AHJ process. Strict evidence for regulatory conclusions; clearly labeled research leads for unresolved issues.")

with st.sidebar:
    st.warning("⚠️ Pay-As-You-Go Active. Successful research results cached for 1 hour; failed calls are never cached.")

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

researcher_notes = st.text_area("Researcher / AHJ notes (optional)", height=160, help="Paste reported answers, open questions, contact/date details and any clarified scope. These remain reported guidance until independently supported.")

st.header("3. Research Execution")
st.caption("One research generation per run. No automatic research, repair or retry calls.")
input_string = f"{PROMPT_VERSION}|{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}|{researcher_notes}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    st.session_state.error_msg = None
    st.session_state.report_data = None
    st.session_state.run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
    st.session_state.run_started_utc = datetime.now(timezone.utc).isoformat()
    st.session_state.run_status = "Started — validating project inputs"
    if not GEMINI_KEY:
        st.session_state.error_msg = "GEMINI_KEY missing."
    else:
        with st.spinner("Researching sources and preparing focused AHJ questions..."):
            prompt = build_single_pass_prompt(address, state, project_date, ptype, bclass, existing_permit, sow_text, researcher_notes)
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
            final_data, final_debug, final_error = finalize_single_pass_result(
                result, address, state, project_date, sow_text, researcher_notes, existing_permit)
            st.session_state.debug_log = final_debug
            st.session_state.report_data = final_data
            st.session_state.error_msg = final_error
            st.session_state.run_status = ("Complete — researcher brief ready" if final_data else "Stopped — no automatic retry made")

with st.expander("Replay a saved research snapshot — no Gemini call"):
    saved_snapshot = st.file_uploader("Research snapshot JSON", type=["json"])
    if saved_snapshot is not None:
        replay_project_date = st.date_input("Original permit / construction date (needed for older snapshots)", value=None)
        if st.button("Rebuild saved brief without an AI request"):
            try:
                replay_data, replay_debug, replay_error = replay_saved_snapshot(json.loads(saved_snapshot.getvalue()), replay_project_date)
                st.session_state.report_data = replay_data
                st.session_state.debug_log = replay_debug
                st.session_state.error_msg = replay_error
                st.session_state.run_status = "Complete — saved response replayed without a Gemini request"
            except Exception as exc:
                st.error(str(exc))

st.info(f"Research status: {st.session_state.run_status}")
if st.session_state.run_id:
    st.caption(f"Run ID: {st.session_state.run_id} · started {st.session_state.run_started_utc}")

if isinstance(st.session_state.debug_log, dict) and st.session_state.debug_log.get("response_warning"):
    st.warning(st.session_state.debug_log["response_warning"])
if st.session_state.error_msg:
    st.error(f"❌ {st.session_state.error_msg}")

validation_errors = st.session_state.debug_log.get("validation_errors", []) if isinstance(st.session_state.debug_log, dict) else []
repair_errors = st.session_state.debug_log.get("validation_repair", {}).get("validation_errors", []) if isinstance(st.session_state.debug_log, dict) else []
visible_errors = repair_errors or validation_errors

if visible_errors:
    st.markdown("**Evidence issues retained for review:**")
    with st.expander("⚠️ View Regulatory Consistency Errors", expanded=True):
        for err in visible_errors:
            st.warning(err)

with st.expander("💰 Gemini API Usage", expanded=bool(st.session_state.error_msg)):
    debug = st.session_state.debug_log if isinstance(st.session_state.debug_log, dict) else {}
    if debug.get("replay_mode"):
        st.caption("Replay made no Gemini request. Any token usage shown below belongs to the original saved run.")
    usage_rows = []
    outer_debug = debug
    if isinstance(debug.get("first_attempt"), dict):
        debug = {**debug["first_attempt"], **{k: v for k, v in outer_debug.items() if k != "first_attempt"}}
    for label, block in (("Initial research", debug), ("Jurisdiction recovery", debug.get("jurisdiction_recovery", {})), ("Evidence resolution", debug.get("evidence_resolution", {})), ("Generation retry", debug.get("retry_attempt", {})), ("Permit recovery", debug.get("permit_recovery", {})), ("Process recovery", debug.get("process_recovery", {})), ("Validation repair", debug.get("validation_repair", {}))):
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
    st.download_button("Download research snapshot for local audit", data=json.dumps({
        "version": PROMPT_VERSION, "project_context": st.session_state.debug_log.get("project_context"), "raw_dossier": (st.session_state.debug_log.get("first_attempt") or st.session_state.debug_log).get("raw_dossier"),
        "final_dossier": st.session_state.get("report_data"), "debug": st.session_state.debug_log,
    }, indent=2, default=str), file_name="ahj_research_snapshot.json", mime="application/json")

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

def _fact_source_label(fact):
    if isinstance(fact, dict) and fact.get("provenance_status") == "SUMMARY_UNVERIFIED":
        return "Scope summary — attribution unverified"
    return _pretty_fact_source(fact.get("source", "UNKNOWN") if isinstance(fact, dict) else "UNKNOWN")


def _code_adoption_summary(code):
    record = code.get("adoption_record") or {}
    if not record: return str(code.get("validation_note") or "Code adoption has not been independently checked.")
    timing = "Effective " + record["effective_date"]
    if record.get("mandatory_date"): timing += "; mandatory " + record["mandatory_date"]
    return timing + ". Source checked: " + str(record.get("url") or "")


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
            if re.search(r"no proposition.specific|not (?:yet )?established for this decision|permit consequence|not_established", qn):
                continue
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
        "a accessibility": "an accessibility",
        "a existing": "an existing",
        "a HVAC": "an HVAC",
        "a ADA": "an ADA",
    }

    def clean(value):
        if isinstance(value, str):
            out = value
            for old, new in replacements.items():
                out = re.sub(r"(?<!\w)" + re.escape(old) + r"(?!\w)", new, out, flags=re.IGNORECASE)
            return out
        if isinstance(value, list):
            return [clean(v) for v in value]
        if isinstance(value, dict):
            return {k: copy.deepcopy(v) if k in {"source_excerpt", "source_quote", "scope_quote", "rule_quote", "quote", "name", "title", "_research_context", "researcher_notes_original"} else clean(v) for k, v in value.items()}
        return value

    return clean(data)


def sanitize_reopen_items(data):
    """Keep only useful project-fact triggers in 'Revisit if'.

    Generic regulatory filler such as "a related regulatory review is identified"
    is removed. A reopen item should tell the project team what FACT CHANGE would
    justify revisiting the conclusion, not merely say that regulation exists.
    """
    if not isinstance(data, dict):
        return data

    generic = re.compile(
        r"^(?:a\s+)?(?:related\s+)?regulatory\s+review\s+is\s+identified|"
        r"^(?:a\s+)?(?:related\s+)?review\s+is\s+identified|"
        r"authoritative\s+research\s+(?:identifies|identified)\s+(?:a\s+)?(?:related\s+)?review",
        re.I,
    )
    # These terms usually describe a legal/process outcome rather than a fact
    # change. Keep them only when the sentence also contains a concrete project
    # condition (e.g. sewer work entering the right-of-way).
    regulatory = re.compile(r"\b(?:permit|approval|clearance|sign[- ]?off|required|must|shall)\b", re.I)
    fact_signal = re.compile(
        r"\b(?:if|when|changes?|extends?|increases?|decreases?|adds?|removes?|relocates?|"
        r"weight|anchorage|tonnage|amperage|valuation|occupancy|right[- ]of[- ]way|"
        r"parking|equipment|service|lateral|wall|roof|structure|location|count|area|scope)\b",
        re.I,
    )

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        raw = item.get("reopen", [])
        if isinstance(raw, str):
            raw = [raw]
        cleaned = []
        for value in raw or []:
            text = str(value or "").strip()
            if not text or generic.search(text):
                continue
            # Preserve the useful sewer/right-of-way fact trigger without
            # asserting the downstream permit consequence.
            lowered = text.lower()
            if ("sewer" in lowered or "lateral" in lowered) and ("right-of-way" in lowered or "right of way" in lowered):
                text = "Sewer work extends into or affects the public right-of-way."
            elif regulatory.search(text) and not fact_signal.search(text):
                continue
            if text not in cleaned:
                cleaned.append(text[:300])
        item["reopen"] = cleaned[:5]
    return data

def sanitize_potential_issues(data):
    """Keep AI research leads structurally separate from regulatory conclusions."""
    if not isinstance(data, dict):
        return data

    allowed_basis = {
        "SOW_FACT",
        "MISSING_FACT",
        "REGULATORY_TOPIC",
        "UNRESOLVED_RELATIONSHIP",
    }
    definitive = re.compile(
        r"\b(?:requires?|required|must|shall|prohibits?|prohibited|exempts?|exempt|"
        r"triggers?|triggered|mandated|is required|are required|is exempt|are exempt|"
        r"permit is|permit required|approval required|clearance required)\b",
        re.I,
    )
    numeric_rule = re.compile(
        r"(?:\b\d+(?:\.\d+)?\s*(?:sq\.?\s*ft|square feet|days?|hours?|feet|ft|inches|in\.?|"
        r"amps?|amperes?|tons?|hp|volts?|kw|kva|percent|%|lb|lbs)\b|"
        r"\b(?:section|sec\.?|chapter|article)\s*\d+[A-Za-z0-9.\-()]*|"
        r"\b(?:title|rule|ordinance|code)\s*\d+[A-Za-z0-9.\-()]*|"
        r"\b\d{2,}\s*(?:day|days|sq|sf|ft|amp|amps|volt|volts|kw|hp)\b)",
        re.I,
    )

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue
        raw = item.get("potential_issues", [])
        if isinstance(raw, str):
            raw = [raw]
        cleaned = []
        for entry in raw or []:
            if isinstance(entry, dict):
                issue = _soften_inference_lead(str(entry.get("issue", "")).strip())
                why = str(entry.get("why", "")).strip()
                basis = str(entry.get("basis", "")).upper().strip()
                if not issue or basis not in allowed_basis:
                    continue

                # A lead is a research hypothesis, not a second evidence channel.
                # Reject it if Gemini smuggled a legal consequence or precise rule
                # into the lead. The human can still see the useful research topic
                # through the discipline's questions/evidence.
                if definitive.search(issue) or definitive.search(why):
                    continue
                if numeric_rule.search(issue) or numeric_rule.search(why):
                    continue

                ev_ids = entry.get("evidence", [])
                if isinstance(ev_ids, str):
                    ev_ids = [ev_ids]
                evidence_by_id = {
                    str(e.get("id")): e for e in (data.get("evidence") or [])
                    if isinstance(e, dict) and e.get("id")
                }

                # A lead's evidence links must point to evidence that is still
                # substantively usable. Merely naming an evidence ID is not enough:
                # quarantined/OTHER evidence cannot legitimize a regulatory or
                # entitlement hypothesis.
                valid_ids = []
                for raw_id in (ev_ids or []):
                    eid = str(raw_id)
                    ev = evidence_by_id.get(eid)
                    if not ev:
                        continue
                    ptype = str(ev.get("proposition_type") or "").upper()
                    if ptype == "OTHER":
                        continue
                    if evidence_proposition_integrity_errors(ev):
                        continue
                    if ptype in {"PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "PATHWAY", "REVIEW_REQUIREMENT", "ENTITLEMENT"}:
                        if evidence_source_specificity_errors([ev]):
                            continue
                    valid_ids.append(eid)
                ev_ids = list(dict.fromkeys(valid_ids))

                # Regulatory-topic / relationship leads need a real, usable
                # evidence anchor. For an unresolved entitlement relationship
                # (for example an existing CUP), require entitlement or
                # applicability evidence specifically; a generic code-current
                # record is not enough. SOW/missing-fact leads may stand on the
                # stated project information alone.
                if basis in {"REGULATORY_TOPIC", "UNRESOLVED_RELATIONSHIP"}:
                    if not ev_ids:
                        continue
                    if basis == "UNRESOLVED_RELATIONSHIP":
                        usable_types = {
                            str(evidence_by_id[eid].get("proposition_type") or "").upper()
                            for eid in ev_ids if eid in evidence_by_id
                        }
                        if not usable_types.intersection({"ENTITLEMENT", "APPLICABILITY", "AUTHORITY_HIERARCHY"}):
                            continue

                regulatory_language = re.compile(r"\b(?:under|pursuant to|per|code|ordinance|municipal code|plan check|permit|zoning|setback|noise standard|life safety|accessibility standard|entitlement|cup|conditional use permit)\b", re.I)
                if not ev_ids and regulatory_language.search(issue + " " + why):
                    continue

                if not why:
                    why = "This is a research lead based on the stated scope; confirm the applicable AHJ rule or project-specific condition."

                cleaned.append({
                    "issue": issue[:400],
                    "why": why[:400],
                    "basis": basis,
                    "evidence": ev_ids,
                })
            else:
                # Legacy/string-only leads are intentionally dropped. Requiring the
                # structured basis prevents unsupported prose from surviving cleanup.
                continue

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
        pathway_line = "No authoritative source found in this research run established the review/submission pathway for this scope."
    if determination == "applies":
        apply_line = "This discipline applies to the stated scope."
    elif determination == "does_not_apply":
        apply_line = "This discipline does not apply to the stated scope."
    else:
        apply_line = "Whether this discipline applies still needs to be confirmed."
    return apply_line, permit_line, pathway_line



def prepare_research_directions(data):
    return prepare_decision_brief(data)


def _discipline_questions(item):
    return ["Decision: " + t["decision"] + "\nAsk " + t["who"] + ": " + t["question"] + "\nWhat the answer changes: " + t["answer_changes"]
        for t in item.get("decision_followups", []) if isinstance(t, dict)]

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


def _set_repeat_table_header(row):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    trPr = row._tr.get_or_add_trPr()
    tbl_header = trPr.find(qn("w:tblHeader"))
    if tbl_header is None:
        tbl_header = OxmlElement("w:tblHeader")
        trPr.append(tbl_header)
    tbl_header.set(qn("w:val"), "true")


def _set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def _set_table_widths(table, widths):
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            cell.width = Inches(width)
            _set_cell_margins(cell)


def _add_label_value(doc, label, value, *, bullet=False):
    value = str(value or "").strip()
    if not value:
        return None
    p = doc.add_paragraph(style="List Bullet" if bullet else None)
    p.paragraph_format.space_after = Pt(2)
    p.add_run(f"{label}: ").bold = True
    p.add_run(value)
    return p


def _add_bullets(doc, values, *, bold_prefix=None):
    if isinstance(values, str):
        values = [values]
    for value in values or []:
        text = str(value or "").strip()
        if not text:
            continue
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        if bold_prefix:
            p.add_run(bold_prefix).bold = True
        p.add_run(text)


def _configure_word_styles(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(9.5)
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.05
    for name, size, before, after in (
        ("Title", 22, 0, 5),
        ("Subtitle", 10.5, 0, 10),
        ("Heading 1", 15, 12, 5),
        ("Heading 2", 12, 9, 3),
        ("Heading 3", 10, 6, 2),
    ):
        style = styles[name]
        style.font.name = "Aptos Display" if name != "Subtitle" else "Aptos"
        style.font.size = Pt(size)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
    styles["List Bullet"].font.name = "Aptos"
    styles["List Bullet"].font.size = Pt(9.5)
    styles["List Bullet"].paragraph_format.space_after = Pt(1)


def _add_source_bullets(doc, evidence_ids, ev_dict):
    seen = set()
    for eid in evidence_ids or []:
        eid = str(eid)
        if eid in seen or eid not in ev_dict:
            continue
        seen.add(eid)
        ev = ev_dict[eid]
        title = str(ev.get("title") or eid).strip()
        rule = str(ev.get("rule") or "").strip()
        url = str(ev.get("url") or "").strip()
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(1)
        p.add_run(title).bold = True
        if rule:
            p.add_run(f" — {rule}")
        if url:
            p2 = doc.add_paragraph(url)
            p2.paragraph_format.left_indent = Inches(0.25)
            p2.paragraph_format.space_after = Pt(2)
            for run in p2.runs:
                run.font.size = Pt(8)


def _discipline_source_ids(item):
    app = item.get("applicability") or {}
    ids = []
    for key, obj in (
        ("evidence", app),
        ("permit_evidence", item),
        ("pathway_evidence", item),
        ("process_evidence", item),
    ):
        raw = obj.get(key, []) if isinstance(obj, dict) else []
        if isinstance(raw, str):
            raw = [raw]
        ids.extend(str(v) for v in (raw or []) if str(v).strip())
    return list(dict.fromkeys(ids))

if st.session_state.report_data:
    report_context = st.session_state.report_data.get("_research_context") or {}
    address = report_context.get("address") or address
    state = report_context.get("state") or state
    project_date = report_context.get("project_date") or project_date
    data = sanitize_generated_prose(st.session_state.report_data)
    data = sanitize_scope_relevance(data)
    data = sanitize_unsupported_applicability(data)
    data = sanitize_unsubstantiated_conditional_pathways(data)
    data = sanitize_expected_process(data)
    data = derive_permit_statuses_from_evidence(data)
    data = sanitize_reopen_items(data)
    # Re-derive process truth after expected-process sanitization so the UI cannot
    # overwrite an evidence-established process finding merely because the model
    # did not populate the optional expected_process prose.
    data = derive_process_status_from_evidence(data)
    st.session_state.report_data = data
    data = sanitize_generic_actionable_questions(st.session_state.report_data)
    data = sanitize_potential_issues(st.session_state.report_data)
    data = sanitize_project_fact_provenance(data)
    data = sanitize_display_claims(data)
    data = prepare_research_directions(data)
    data = build_validated_bottom_line(data)
    st.session_state.report_data = data
    st.divider()
    st.header("4. Research Dossier")
    if "validation_errors" in st.session_state.debug_log:
        st.warning("Some source-supported findings could not be published. The conservative research brief and diagnostics remain available.")
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
    if data.get("scope_coverage_gaps"):
        st.subheader("Scope topics still needing research")
        st.caption("These are coverage gaps, not permit requirements. No additional AI call was made.")
        for gap in data["scope_coverage_gaps"]:
            st.markdown("**" + gap["topic"] + "** — " + gap["scope_quote"])
    if data.get("project_research_tasks"):
        st.subheader("Shared research tasks")
        for task in data["project_research_tasks"]:
            st.markdown("- " + task)
    if data.get("researcher_notes_original"):
        with st.expander("Original researcher / AHJ notes"):
            st.text(data["researcher_notes_original"])
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
        st.caption(_code_adoption_summary(code))
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
                st.metric("Scope signal", _scope_display(item))
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
            for excerpt in item.get("scope_excerpts", []):
                st.markdown("**Supplied scope:** " + excerpt)
            st.markdown("**Research direction:** " + str((item.get("research_direction") or {}).get("summary") or ""))
            for entry in item.get("reported_guidance", []):
                st.markdown("**" + entry["label"] + ":** " + entry["source_quote"])
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
                st.markdown(f"**Project information:** {fact_statement} [{_fact_source_label(fact)}]")
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
        _configure_word_styles(doc)

        _add_doc_title(doc, "AHJ Research Dossier", f"{address} · {state} · {project_date}")
        meta = doc.add_paragraph()
        meta.paragraph_format.space_after = Pt(10)
        meta.add_run(f"Generated {datetime.now().strftime('%B %d, %Y')}  |  Research assistant output - verify unresolved items with the AHJ.").italic = True

        # Executive research brief: designed to be useful even when most legal outcomes remain unresolved.
        doc.add_heading("Research Brief", level=1)
        if data.get("_generation_incomplete"):
            doc.add_paragraph("Partial research: the provider ended generation early. Only complete returned data was evaluated; this brief may omit disciplines or issues.")
        bottom = str(data.get("bottom_line") or "").strip()
        if bottom:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(7)
            p.add_run(bottom).bold = True

        completeness = data.get("research_completeness") or {}
        critical_missing = completeness.get("critical_missing", []) if isinstance(completeness, dict) else []
        if isinstance(critical_missing, str):
            critical_missing = [critical_missing]
        if critical_missing:
            doc.add_heading("Best next information to obtain", level=2)
            _add_bullets(doc, critical_missing)

        if data.get("scope_coverage_gaps"):
            doc.add_heading("Scope topics still needing research", level=2)
            doc.add_paragraph("These supplied work topics have no matched finding, reported answer or retained decision question. Review them with the project team/researcher. They do not establish permit requirements.")
            _add_bullets(doc, [g["topic"] + ": " + g["scope_quote"] for g in data["scope_coverage_gaps"]])
        if data.get("project_research_tasks"):
            doc.add_heading("Shared research tasks", level=2)
            _add_bullets(doc, data["project_research_tasks"])
        doc.add_heading("Jurisdiction", level=1)
        table = doc.add_table(rows=4, cols=2)
        table.style = "Light Shading Accent 1"
        _set_table_widths(table, [1.45, 5.65])
        rows = [
            ("Status", _pretty_status(jur.get("status", "N/A"))),
            ("County", jur.get("county", "N/A")),
            ("City", jur.get("city", "N/A")),
            ("AHJ", jur.get("ahj", "N/A")),
        ]
        for row, (label, value) in zip(table.rows, rows):
            row.cells[0].text = label
            row.cells[1].text = str(value)
            row.cells[0].paragraphs[0].runs[0].bold = True
            _shade_cell(row.cells[0], "EAF2F8")

        doc.add_heading("Applicable Codes", level=1)
        for code in (data.get("codes") or []):
            name = str(code.get("name") or "Unknown code")
            status = _pretty_status(code.get("status"))
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(1)
            p.add_run(name).bold = True
            p.add_run(f" — {status}")
            doc.add_paragraph(_code_adoption_summary(code))

        doc.add_heading("Permit & Review Snapshot", level=1)
        p = doc.add_paragraph("This table separates scope relevance from permit and review conclusions. 'Not yet established' is intentionally unresolved, not a negative finding.")
        p.paragraph_format.space_after = Pt(5)
        matrix = doc.add_table(rows=1, cols=4)
        matrix.style = "Light Shading Accent 1"
        _set_table_widths(matrix, [2.25, 1.25, 1.75, 1.85])
        hdr = matrix.rows[0].cells
        for cell, text in zip(hdr, ["Discipline", "Scope", "Permit", "Review / pathway"]):
            cell.text = text
            _shade_cell(cell, "D9EAF7")
            cell.paragraphs[0].runs[0].bold = True
        _set_repeat_table_header(matrix.rows[0])
        for item in data.get("disciplines", []):
            row = matrix.add_row().cells
            row[0].text = str(item.get("type", "Unknown"))
            row[1].text = _scope_display(item)
            row[2].text = _pretty_status(item.get("permit"))
            row[3].text = _pretty_pathway_status(item.get("pathway"))

        doc.add_page_break()
        doc.add_heading("Discipline Research", level=1)
        intro = doc.add_paragraph("Each section is intentionally compact: scope signal, what is actually known, questions that move the research forward, and useful sources/leads.")
        intro.paragraph_format.space_after = Pt(8)

        for idx, item in enumerate(data.get("disciplines", []), start=1):
            discipline = str(item.get("type") or "Unknown")
            doc.add_heading(f"{idx}. {discipline}", level=2)

            # Compact status card instead of repeated generic prose.
            status_table = doc.add_table(rows=2, cols=3)
            status_table.style = "Light Shading Accent 1"
            _set_table_widths(status_table, [2.35, 2.35, 2.35])
            for cell, text in zip(status_table.rows[0].cells, ["Scope signal", "Permit", "Review / pathway"]):
                cell.text = text
                _shade_cell(cell, "EAF2F8")
                cell.paragraphs[0].runs[0].bold = True
            for cell, text in zip(status_table.rows[1].cells, [
                _scope_display(item),
                _pretty_status(item.get("permit")),
                _pretty_pathway_status(item.get("pathway")),
            ]):
                cell.text = text

            scope_basis = str(item.get("scope_basis") or "").strip()
            if scope_basis:
                _add_label_value(doc, "Why it is in the research", scope_basis, bullet=True)

            app = item.get("applicability") or {}
            fact = app.get("fact", {})
            fact_statement = str(fact.get("statement") or "").strip() if isinstance(fact, dict) else str(fact or "").strip()
            fact_source = fact.get("source", "UNKNOWN") if isinstance(fact, dict) else "UNKNOWN"
            if fact_statement:
                _add_label_value(doc, "Project fact", f"{fact_statement} [{_fact_source_label(fact)}]", bullet=True)

            # Only show substantive established findings; suppress repetitive generic unresolved sentences.
            permit_status = str(item.get("permit") or "UNKNOWN").upper()
            pathway_status = str(item.get("pathway") or "UNKNOWN").upper()
            substantive = []
            if permit_status in {"VERIFIED_REQUIRED", "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED"}:
                finding = str(item.get("permit_finding") or "").strip()
                if finding:
                    substantive.append(("Permit finding", finding))
            if pathway_status in {"VERIFIED_REQUIRED", "NOT_APPLICABLE"}:
                finding = str(item.get("pathway_finding") or "").strip()
                if finding:
                    substantive.append(("Review finding", finding))
            process_finding = str(item.get("process_finding") or "").strip()
            expected_process = str(item.get("expected_process") or "").strip()
            process_ids = item.get("process_evidence") or []
            # Only place process language under "What the research established" when
            # proposition-specific process evidence actually survived validation.
            # Otherwise the status card already says "Not yet established" and we
            # avoid repeating a generic sentence in every discipline.
            if process_finding and process_ids:
                substantive.append(("Process evidence", process_finding))
            if expected_process and process_ids:
                substantive.append(("AHJ process", expected_process))
            if substantive:
                doc.add_heading("What the research established", level=3)
                for label, value in substantive:
                    _add_label_value(doc, label, value, bullet=True)

            process_details = [
                ("Who reviews", item.get("process_reviewer")),
                ("How it moves", item.get("process_route")),
                ("Relationship", item.get("process_relationship")),
                ("What changes it", item.get("process_trigger")),
            ]
            process_details = [(a, str(b or "").strip()) for a, b in process_details if str(b or "").strip()]
            if process_details:
                doc.add_heading("Process details", level=3)
                for label, value in process_details:
                    _add_label_value(doc, label, value, bullet=True)

            if item.get("scope_excerpts"):
                doc.add_heading("Relevant supplied scope", level=3)
                _add_bullets(doc, item["scope_excerpts"])
            direction = str((item.get("research_direction") or {}).get("summary") or "")
            if direction:
                _add_label_value(doc, "Research direction", direction)
            if item.get("reported_guidance"):
                doc.add_heading("Reported AHJ guidance", level=3)
                for entry in item["reported_guidance"]:
                    _add_label_value(doc, entry["label"], entry["source_quote"])
            questions = _discipline_questions(item)
            if questions:
                doc.add_heading("Questions that move this forward", level=3)
                _add_bullets(doc, questions)

            missing = item.get("missing", [])
            if isinstance(missing, str):
                missing = [missing]
            missing = [str(v).strip() for v in (missing or []) if str(v).strip()]
            if missing:
                doc.add_heading("Information still needed", level=3)
                _add_bullets(doc, missing)

            potential_issues = item.get("potential_issues") or []
            if potential_issues:
                doc.add_heading("Worth checking - AI-generated possible lead", level=3)
                for lead in potential_issues:
                    if not isinstance(lead, dict):
                        continue
                    issue = str(lead.get("issue") or "").strip()
                    why = str(lead.get("why") or "").strip()
                    if issue:
                        p = doc.add_paragraph(style="List Bullet")
                        p.paragraph_format.space_after = Pt(1)
                        p.add_run(issue).bold = True
                    if why:
                        p = doc.add_paragraph()
                        p.paragraph_format.left_indent = Inches(0.25)
                        p.paragraph_format.space_after = Pt(2)
                        p.add_run("Basis: ").bold = True
                        p.add_run(why)

            reopen = item.get("reopen", [])
            if isinstance(reopen, str):
                reopen = [reopen]
            reopen = [str(v).strip() for v in (reopen or []) if str(v).strip()]
            if reopen:
                doc.add_heading("Revisit if project facts change", level=3)
                _add_bullets(doc, reopen)

            source_ids = _discipline_source_ids(item)
            if source_ids:
                doc.add_heading("Useful sources", level=3)
                _add_source_bullets(doc, source_ids, ev_dict)

            # Visual separation without wasting a full page per discipline.
            if idx < len(data.get("disciplines", [])):
                p = doc.add_paragraph("─" * 70)
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.size = Pt(6)

        if data.get("researcher_notes_original"):
            doc.add_heading("Original researcher notes", level=1)
            doc.add_paragraph("Supplied conversation notes; not independently verified legal requirements. Scope clarified during those conversations may extend beyond the starting SOW.")
            doc.add_paragraph(data["researcher_notes_original"])

        # Source appendix preserves the evidence trail without cluttering each unresolved section.
        evidence = [ev for ev in (data.get("evidence") or []) if isinstance(ev, dict)]
        if evidence:
            doc.add_page_break()
            doc.add_heading("Evidence Appendix", level=1)
            doc.add_paragraph("Authoritative and supporting sources retained by the research engine. Evidence type matters: a source can establish code currency or applicability without establishing a permit requirement.")
            for ev in evidence:
                eid = str(ev.get("id") or "").strip()
                title = str(ev.get("title") or eid or "Source").strip()
                p = doc.add_paragraph(style="List Bullet")
                p.paragraph_format.space_after = Pt(1)
                p.add_run(f"{eid} - {title}").bold = True
                ptype = str(ev.get("proposition_type") or "OTHER").strip()
                if ptype:
                    p.add_run(f" [{ptype}]")
                if ptype == "CODE_CURRENCY" and ev.get("source_verification") != "SOURCE_CHECKED":
                    doc.add_paragraph("Unverified adoption extraction. The generated summary below has not been matched to a dated official source record.")
                if ptype == "OTHER":
                    doc.add_paragraph("Unverified research lead. The following generated source summary is retained for review and does not establish a requirement.")
                rule = str(ev.get("rule") or "").strip()
                if rule:
                    p2 = doc.add_paragraph(rule)
                    p2.paragraph_format.left_indent = Inches(0.25)
                    p2.paragraph_format.space_after = Pt(1)
                url = str(ev.get("url") or "").strip()
                if url:
                    p3 = doc.add_paragraph(url)
                    p3.paragraph_format.left_indent = Inches(0.25)
                    p3.paragraph_format.space_after = Pt(3)
                    for run in p3.runs:
                        run.font.size = Pt(8)

        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button(
            "📄 Download Word Research Brief",
            data=buf.getvalue(),
            file_name="AHJ_Research_Brief.docx",
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
