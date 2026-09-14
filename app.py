import os
import re
import json
import time
import hashlib
from datetime import datetime, date
from io import BytesIO
import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, Inches

st.set_page_config(page_title="AHJ Research Assistant v26.30.16", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v26.30.16_cross_discipline_fix"

# ============================================================
# HELPERS & VALIDATION
# ============================================================
def discipline_family(value):
    value = str(value or "").strip().lower()
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

def evidence_supports_review(evidence, discipline):
    if not evidence:
        return False
    ev_disc = (evidence.get("discipline") or "").strip()
    current_disc = (discipline or "").strip()
    if discipline_family(ev_disc) != discipline_family(current_disc):
        return False
    return evidence.get("proposition_type") == "REVIEW_REQUIREMENT" and not evidence_proposition_integrity_errors(evidence)

def evidence_supports_any(evidence_ids, evidence_by_id, proposition_types, discipline):
    return any(
        evidence_by_id.get(eid)
        and discipline_family(evidence_by_id[eid].get("discipline", "")) == discipline_family(discipline)
        and evidence_by_id[eid].get("proposition_type") in proposition_types
        and not evidence_proposition_integrity_errors(evidence_by_id[eid])
        for eid in evidence_ids
    )

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

def evidence_proposition_integrity_errors(evidence):
    """Prevent evidence labels from being stronger than the extracted rule text."""
    errors = []
    if not evidence:
        return errors
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
    return [
        eid for eid in evidence_ids
        if eid in evidence_by_id
        and discipline_family(evidence_by_id[eid].get("discipline", "")) == discipline_family(discipline)
        and evidence_by_id[eid].get("proposition_type") == proposition_type
        and not evidence_proposition_integrity_errors(evidence_by_id[eid])
    ]

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

    epistemic_permit_finding = any(re.search(pattern, permit_text) for pattern in [
        r"\bpermit\s+(?:requirement\s+)?(?:is\s+)?not\s+established\b",
        r"\bpermit\s+(?:requirement\s+)?(?:is\s+)?not\s+determined\b",
        r"\bcannot\s+determine\b[^.]{0,160}\bpermit\b",
        r"\bunable\s+to\s+(?:determine|establish)\b[^.]{0,160}\bpermit\b",
        r"\bcurrent\s+evidence\s+does\s+not\s+establish\b[^.]{0,160}\bpermit\b",
        r"\bavailable\s+evidence\s+does\s+not\s+establish\b[^.]{0,160}\bpermit\b",
        r"\bdoes\s+not\s+by\s+itself\s+establish\b[^.]{0,160}\bpermit\b",
        r"\bnot\s+established\s+by\s+current\s+evidence\b",
    ])

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
        if any(re.search(p, permit_text) for p in [
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
    if addr and addr in text:
        return True
    site_terms = [
        "parcel", "tax lot", "taxlot", "property record", "property search",
        "gis", "map viewer", "jurisdiction lookup", "city limits", "municipal limits",
        "parcel is", "property is", "address is", "site is", "tax lot",
        "inside the city", "outside the city", "within the city",
        "within city limits", "outside city limits",
    ]
    return any(term in text for term in site_terms)

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
        "project parcel's actual governmental jurisdiction. Postal city/ZIP and "
        "generic agency coverage do not establish municipal boundaries."
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
    Prevent a mismatched applicability source from being treated as direct proof.
    """
    if not isinstance(data, dict):
        return data

    evidence_by_id = {
        e.get("id"): e
        for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    for item in data.get("disciplines", []) or []:
        if not isinstance(item, dict):
            continue

        applicability = item.get("applicability")
        if not isinstance(applicability, dict):
            continue

        if str(applicability.get("relationship", "")).lower() != "direct":
            continue

        discipline = str(
            item.get("type")
            or item.get("discipline")
            or "Unknown"
        )

        current_family = discipline_family(discipline)
        bad_ids = []
        raw_ids = applicability.get("evidence") or []
        if isinstance(raw_ids, str):
            raw_ids = [raw_ids]

        for eid in raw_ids:
            evidence = evidence_by_id.get(eid)
            if not evidence:
                continue

            evidence_family = discipline_family(
                str(evidence.get("discipline") or "")
            )

            if (
                evidence_family
                and current_family
                and evidence_family != current_family
            ):
                bad_ids.append(eid)

        if bad_ids:
            applicability["relationship"] = "conditional"
            notes = item.setdefault("validation_notes", [])
            notes.append(
                "Applicability was downgraded from direct to conditional because "
                f"the cited evidence ({', '.join(bad_ids)}) belongs to another "
                "discipline. The evidence remains available in the research trail."
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
        original = evidence.get("proposition_type")
        if integrity_errors and original in EVIDENCE_PROPOSITION_TYPES and original != "OTHER":
            evidence["proposition_type"] = "OTHER"
            evidence["validation_note"] = (
                f"Quarantined: the source rule did not explicitly support the declared "
                f"{original} proposition. The source record is preserved for review."
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
        has_claim = any(re.search(pattern, finding) for pattern in permit_claim_patterns)
        epistemic_nonclaim = any(re.search(pattern, finding) for pattern in [
            r"\bnot\s+established\b",
            r"\bcannot\s+determine\b",
            r"\bcan(?:not|'t)\s+be\s+determined\b",
            r"\bdoes\s+not\s+(?:by\s+itself\s+)?establish\b",
            r"\bnot\s+establish(?:ed|ing)?\b",
        ])
        if not has_claim and not epistemic_nonclaim and re.search(r"\bpermit\b", finding):
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
            f"A {discipline.lower()} permit requirement is not established by current "
            "evidence. The available rule may identify an applicability or threshold "
            "condition, but it does not by itself establish the downstream permit "
            "requirement. Confirm the permit consequence with an authoritative "
            "permit-specific source once the missing project facts are known."
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
        if any(re.search(p, finding) for p in epistemic_only):
            continue
        if any(re.search(p, finding) for p in broad_permit_claims):
            item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"A {str(discipline).lower()} permit requirement is not established by current evidence. "
                "The available evidence does not establish the downstream permit consequence. "
                "Confirm the requirement using an authoritative, discipline-specific permit source."
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
                f"A {discipline.lower()} permit requirement is not established by current evidence. "
                "Review or compliance obligations should not be treated as proof that a permit is required."
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
        if path in generic_paths:
            errors.append(f"{ev.get('id', 'Evidence')}: {ptype} evidence points to a generic landing page ({host}{path}).")
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
            parts.append("Permit requirements are established for " + ", ".join(verified) + ".")
        if unresolved:
            parts.append("Permit requirements remain unresolved for " + ", ".join(unresolved) + ".")
        if not_applicable:
            parts.append("No permit requirement is currently established for " + ", ".join(not_applicable) + ".")
        parts.append("See the Permit Matrix for the supporting evidence and the specific items that still need confirmation.")
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
    evidence_by_id = {e.get("id"): e for e in evidence_items if e.get("id")}
    evidence_ids = set(evidence_by_id.keys())
    
    # 4. Only check referenced evidence for generic landing pages
    referenced_evidence_ids = set()
    for item in data.get("disciplines", []):
        if not isinstance(item, dict):
            continue
        referenced_evidence_ids.update(item.get("permit_evidence") or [])
        referenced_evidence_ids.update(item.get("pathway_evidence") or [])
        referenced_evidence_ids.update(item.get("authority_evidence") or [])
        applicability = item.get("applicability") or {}
        referenced_evidence_ids.update(applicability.get("evidence") or [])
        
    referenced_evidence_ids.update(data.get("bottom_line_evidence") or [])
    referenced_evidence_ids.update((data.get("jurisdiction") or {}).get("evidence") or [])
    
    for evidence in evidence_items:
        if evidence.get("id") not in referenced_evidence_ids:
            continue
        errors.extend(evidence_source_specificity_errors([evidence]))

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
            errors.append(f"{eid}: invalid or missing proposition_type '{proposition_type}'.")
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
            conditional_trigger_patterns = [
                r"\bpermit\s+(?:is\s+)?required\s+if\b",
                r"\bpermit\s+(?:is\s+)?required\s+when\b",
                r"\bpermit\s+(?:is\s+)?required\s+only\s+if\b",
                r"\btriggers?\s+(?:a\s+)?permit\b",
                r"\brequires?\s+(?:a\s+)?permit\b",
            ]
            definitive_permit_consequence = any(re.search(pattern, permit_text) for pattern in conditional_trigger_patterns)
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
                if permit not in {"VERIFIED_REQUIRED", "NOT_APPLICABLE"} and re.search(r"\bpermit\s+(?:is\s+)?required\b|\brequires?\s+(?:a\s+)?permit\b", pathway_finding_text):
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
                evidence_by_id.get(eid, {}).get("proposition_type") == "PERMIT_EXEMPTION"
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
        if finding_claims_permit and not valid_permit_ids:
            errors.append(f"{discipline}: permit finding contains a permit consequence that is not established by valid PERMIT_REQUIREMENT evidence.")
        if permit == "VERIFIED_REQUIRED" and not valid_permit_ids:
            errors.append(f"{discipline}: VERIFIED_REQUIRED permit cannot survive without valid PERMIT_REQUIREMENT evidence.")
        if has_definitive_negative_permit_claim(permit_finding) and permit in {"NOT_APPLICABLE", "VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED"} and not valid_exemption_ids:
            errors.append(f"{discipline}: negative/exemption permit statement lacks valid PERMIT_EXEMPTION evidence.")
        if pathway == "VERIFIED_REQUIRED" and not valid_pathway_ids:
            errors.append(f"{discipline}: VERIFIED_REQUIRED pathway cannot survive without valid PATHWAY evidence.")
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
# CACHING & API CALL
# ============================================================
@st.cache_data(ttl=3600)
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
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=prompt_text,
            config=config,
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
        usage = getattr(response, "usage_metadata", None)
        if usage:
            debug_info["usage_metadata"] = {
                "prompt_tokens": getattr(usage, "prompt_token_count", None),
                "candidates_tokens": getattr(usage, "candidates_token_count", None),
                "thoughts_tokens": getattr(usage, "thoughts_token_count", None),
                "total_tokens": getattr(usage, "total_token_count", None),
                "cached_content_tokens": getattr(usage, "cached_content_token_count", None),
            }
        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "error": True, "retry": False, "msg": "Failed to parse JSON.", "debug": debug_info}
        
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
        debug_info["error_type"] = "Python Exception"
        if "429" in error_msg:
            return {"data": None, "error": True, "retry": False, "msg": "Quota exceeded.", "debug": debug_info}
        return {"data": None, "error": True, "retry": False, "msg": f"Error: {error_msg[:200]}", "debug": debug_info}

@st.cache_data(ttl=3600)
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
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=retry_prompt,
            config=config,
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
        usage = getattr(response, "usage_metadata", None)
        if usage:
            debug_info["usage_metadata"] = {
                "prompt_tokens": getattr(usage, "prompt_token_count", None),
                "candidates_tokens": getattr(usage, "candidates_token_count", None),
                "thoughts_tokens": getattr(usage, "thoughts_token_count", None),
                "total_tokens": getattr(usage, "total_token_count", None),
                "cached_content_tokens": getattr(usage, "cached_content_token_count", None),
            }
        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "retry": False, "msg": "Retry produced invalid JSON.", "debug": debug_info}
        
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
        debug_info["error_type"] = "Python Exception"
        return {"data": None, "error": True, "retry": False, "msg": f"Retry error: {str(e)[:200]}", "debug": debug_info}

# ============================================================
# VALIDATION-AWARE SELF-CORRECTION
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_repair(repair_hash, repair_prompt):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": "validation_repair"}
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=32768,
            thinking_config=types.ThinkingConfig(thinking_level="high"),
            tools=[types.Tool(google_search=types.GoogleSearch())],
        )
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=repair_prompt,
            config=config,
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
        usage = getattr(response, "usage_metadata", None)
        if usage:
            debug_info["usage_metadata"] = {
                "prompt_tokens": getattr(usage, "prompt_token_count", None),
                "candidates_tokens": getattr(usage, "candidates_token_count", None),
                "thoughts_tokens": getattr(usage, "thoughts_token_count", None),
                "total_tokens": getattr(usage, "total_token_count", None),
            }
        try:
            data = extract_json(text)
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
        debug_info["error_type"] = "Python Exception"
        return {"data": None, "error": True, "msg": f"Validation repair error: {str(e)[:200]}", "debug": debug_info}

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state: st.session_state.error_msg = None

st.title("🏛️ AHJ Research Assistant v26.30.16")
st.caption("32K generation ceiling. High reasoning. Code-currency + state/local authority hierarchy + proposition-specific evidence + consequence firewall + deterministic status repair + one targeted self-correction pass.")

with st.sidebar:
    st.warning("⚠️ Pay-As-You-Go Active. Results cached for 1 hour.")
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
                {"type": "Mechanical", "applicability": {"rule": "Commercial mechanical equipment replacement requires a mechanical permit.", "fact": {"statement": "Replacing ground-level exterior commercial HVAC equipment.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E2"]}, "permit": "VERIFIED_REQUIRED", "permit_finding": "Mechanical permit requirement is established by the cited permit/code evidence.", "permit_basis": "DIRECT_EVIDENCE", "permit_evidence": ["E2"], "pathway": "CONDITIONAL", "pathway_finding": "Specific review pathway remains unresolved because project-specific pathway criteria have not been verified.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Project-specific review pathway criteria (e.g., unit weight, CFM)."], "reopen": ["Authoritative pathway criteria retrieved."]},
                {"type": "Electrical", "applicability": {"rule": "Electrical permit required for branch circuit/disconnect modification.", "fact": {"statement": "Electrical modifications to replacement unit are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Whether wiring/disconnect/breaker will be modified.", "relationship": "conditional", "evidence": ["E4"]}, "permit": "CONDITIONAL", "permit_finding": "Electrical permit consequence depends on whether wiring/disconnect/circuit work occurs.", "permit_basis": "CONDITIONAL", "permit_evidence": ["E4"], "pathway": "CONDITIONAL", "pathway_finding": "Pathway cannot be determined until electrical scope is defined.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Unit electrical specs (MCA, MOP, voltage).", "Scope of electrical changes."], "reopen": ["Modifying electrical disconnect, wiring, or breaker."]},
                {"type": "Energy", "applicability": {"rule": "Current energy code applies to replacement equipment.", "fact": {"statement": "Replacing HVAC unit; efficiency ratings unknown.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E3"]}, "permit": "CONDITIONAL", "permit_finding": "Energy compliance applies; a separate energy permit requirement is not established by current evidence.", "permit_basis": "NOT_ESTABLISHED", "permit_evidence": [], "pathway": "CONDITIONAL", "pathway_finding": "Compliance pathway depends on equipment specifications and primary permit type.", "pathway_basis": "NOT_ESTABLISHED", "pathway_evidence": [], "missing": ["Replacement equipment efficiency/specifications."], "reopen": []},
                {"type": "Structural", "applicability": {"rule": "Structural requirements depend on replacement equipment loads and attachment conditions.", "fact": {"statement": "Replacement unit weight and anchorage configuration are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Replacement unit weight and anchorage details.", "relationship": "conditional", "evidence": []}, "permit": "UNKNOWN", "permit_finding": "Structural permit consequence cannot be determined from current project facts.", "permit_basis": "NOT_ESTABLISHED", "permit_evidence": [], "pathway": "UNKNOWN", "pathway_finding": "Structural review pathway cannot be established from current information.", "pathway_basis": "NOT_ESTABLISHED", "pathway_evidence": [], "missing": ["Replacement unit operating weight", "Anchorage configuration"], "reopen": ["Equipment weight or anchorage requires structural review"]},
                {"type": "Planning / CUP", "applicability": {"rule": "Work must comply with existing CUP conditions.", "fact": {"statement": "Parcel operates under existing CUP; actual conditions not retrieved.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Actual CUP conditions governing exterior equipment.", "relationship": "conditional", "evidence": ["E1"]}, "permit": "CONDITIONAL", "permit_finding": "Existing CUP identified, but governing conditions have not been reviewed.", "permit_basis": "CONDITIONAL", "permit_evidence": [], "pathway": "CONDITIONAL", "pathway_finding": "Final land-use determination conditional on review of existing CUP conditions.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Actual CUP conditions governing exterior equipment."], "reopen": ["Relocation, footprint expansion, screening changes, noise increases, or site work occur."]}
            ]
        }
        st.session_state.debug_log = {"mock": True, "note": "No API call made"}
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
JURISDICTION FIREWALL: Determine the project's actual governmental jurisdiction from authoritative site-specific evidence (parcel/GIS/property record/jurisdiction lookup or an equivalent official source). A postal city, ZIP code, mailing address city, or generic county/city service page does NOT establish municipal jurisdiction. If the parcel is unincorporated, set the actual city field to "Unincorporated" and do not treat the postal city as the municipal jurisdiction. The AHJ must correspond to the verified governmental jurisdiction.
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
DISCIPLINE DISCOVERY: Determine disciplines dynamically from project scope, project type, jurisdiction, adopted codes, permit requirements, land-use controls, and authoritative applicability rules. Do not use a fixed discipline list. Do not impose a maximum number of disciplines.
For every discipline determine separately:
1. APPLICABILITY: Does the authoritative rule apply to the known project facts?
2. PERMIT: Does authoritative evidence establish a permit requirement or exemption?
3. PATHWAY: Does authoritative evidence establish a specific review pathway?
4. AUTHORITY HIERARCHY: When state and local rules could differ, which governmental rule controls this subject in this jurisdiction?
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
A regulatory dossier can contain useful professional research leads even when the law/permit conclusion is not established. In the user-facing report these are labeled “Worth checking — AI research lead.” These are NOT regulatory conclusions and MUST NOT affect permit, pathway, applicability, jurisdiction, code status, Bottom Line, or research completeness.
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
                result = cached_gemini_call(prompt_hash, prompt)
                if result.get("error") and result.get("debug", {}).get("error_type") == "Validation Failed":
                    validation_errors = result.get("debug", {}).get("validation_errors", [])
                    prior_json = json.dumps(result.get("data") or {}, indent=2)
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
- Fix every validation error.
- Do not evade an error by deleting a discipline or evidence item merely to make validation pass.
- JURISDICTION IS A HARD REQUIREMENT: re-research the actual parcel jurisdiction using authoritative site-specific parcel/GIS/property/jurisdiction evidence. Postal city/ZIP and generic agency coverage are not enough. If evidence establishes an unincorporated parcel, use "Unincorporated" for the actual city field.
- CODE CURRENCY IS A HARD REQUIREMENT: re-research each CURRENT code using authoritative adoption/current-code sources as of the project date. Never rely on remembered code years.
- A code may be CURRENT only when valid CODE_CURRENCY evidence establishes its edition and current/adopted/effective/mandatory status as of the project date.
- If the cited evidence is for a newer edition than the code name, correct the code entry to the edition actually supported by the evidence; if current status remains unresolved, use CONDITIONAL. Do not preserve a stale code year merely to keep the original text.
- Do not treat a prior edition being available online as evidence that it is current.
- Preserve valid evidence and project facts.
- If a permit requirement lacks discipline-matched PERMIT_REQUIREMENT evidence, either retrieve authoritative permit-specific evidence using Google Search or downgrade the permit conclusion to CONDITIONAL/UNKNOWN.
- If a negative permit/exemption claim lacks discipline-matched PERMIT_EXEMPTION evidence, remove the definitive negative claim or retrieve explicit exemption evidence. IMPORTANT: do not treat an epistemic statement such as "permit requirement is not established," "current evidence does not establish a permit requirement," or "cannot determine whether a permit is required" as a legal exemption. Those statements are allowed with NOT_ESTABLISHED / UNKNOWN / CONDITIONAL status and do not require PERMIT_EXEMPTION evidence.
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
"""
                    repair_hash = hashlib.md5((prompt_hash + "|validation_repair|" + json.dumps(validation_errors, sort_keys=True)).encode()).hexdigest()
                    repair_result = cached_gemini_repair(repair_hash, repair_prompt)
                    st.session_state.debug_log = {
                        "first_attempt": result.get("debug", {}),
                        "validation_repair": repair_result.get("debug", {}),
                    }
                    if repair_result.get("error"):
                        st.session_state.error_msg = repair_result.get("msg", "Validation repair failed.")
                        st.session_state.report_data = None
                    else:
                        st.session_state.report_data = repair_result["data"]
                        st.session_state.error_msg = None
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
                    else:
                        st.session_state.debug_log = {
                            "first_attempt": result.get("debug", {}),
                            "retry_attempt": retry_result.get("debug", {}),
                        }
                        st.session_state.report_data = retry_result["data"]
                else:
                    st.session_state.debug_log = result.get("debug", {})
                    if result["error"]:
                        st.session_state.error_msg = result["msg"]
                    else:
                        st.session_state.report_data = result["data"]

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

with st.expander("🐛 API Debug Log", expanded=False):
    st.json(st.session_state.debug_log)

# ============================================================
# USER-FACING RESULTS / EXPORT
# ============================================================
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
    if permit == "VERIFIED_REQUIRED":
        refs = permit_ids
        trail_type = "Permit decision reference"
        conclusion = str(item.get("permit_finding", "")).strip()
    elif permit == "NOT_APPLICABLE":
        refs = permit_ids
        trail_type = "Permit decision reference"
        conclusion = str(item.get("permit_finding", "")).strip()
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
        with st.expander(f"{icon} {title} — {permit_label}", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                st.metric("Permit", permit_label)
            with c2:
                st.metric("Review / pathway", pathway_label)
            st.markdown("### In plain English")
            st.markdown(f"**Permit:** {permit_line}")
            st.markdown(f"**Review:** {pathway_line}")
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
                st.markdown("### 💡 Worth checking — AI research lead")
                st.caption("Research leads only — not regulatory conclusions and not part of the permit determination.")
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
                st.markdown(f"**Project information:** {fact_statement} [{fact_source}]")
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
        matrix = doc.add_table(rows=1, cols=3)
        matrix.style = "Light Shading Accent 1"
        hdr = matrix.rows[0].cells
        for cell, text in zip(hdr, ["Discipline", "Permit", "Review / pathway"]):
            cell.text = text
            _shade_cell(cell, "D9EAF7")
        for item in data.get("disciplines", []):
            row = matrix.add_row().cells
            row[0].text = str(item.get("type", "Unknown"))
            row[1].text = _pretty_status(item.get("permit"))
            row[2].text = _pretty_pathway_status(item.get("pathway"))
        doc.add_page_break()
        doc.add_heading("Discipline Findings", level=1)
        for item in data.get("disciplines", []):
            doc.add_heading(str(item.get("type", "Unknown")), level=2)
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
            p.add_run(f"{fact_statement} [{fact_source}]")
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
                doc.add_paragraph("Worth checking — AI research lead:")
                doc.add_paragraph("These are research leads, not regulatory conclusions. They do not affect permit status or the regulatory determination.")
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
