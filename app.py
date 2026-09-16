# app.py
# AHJ Research Assistant – Test / Development Harness
# Focus: deterministic evidence contract + mock research path
# Safe to run without a Gemini key (Mock Mode default = True)

import os
import re
import json
import hashlib
import copy
from datetime import datetime, date, timezone
from io import BytesIO

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches

# Optional real Gemini path
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

st.set_page_config(
    page_title="AHJ Research Assistant – Test Harness",
    page_icon="🏛️",
    layout="wide",
)

# ============================================================
# CONFIG
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

PROJECT_TYPES = [
    "Replacement / Repair", "Remodel / Tenant Improvement",
    "Addition", "New Construction", "Site / Civil Work", "Other",
]

BUILDING_CLASSES = [
    "Commercial", "Assembly", "Institutional", "Industrial",
    "Agricultural", "Residential (1-2 Family)", "Residential (Multi-family)",
    "Mixed-use", "Unknown",
]

FACT_SOURCES = {"USER_PROVIDED", "RETRIEVED_RECORD", "AUTHORITATIVE_SOURCE", "INFERRED", "UNKNOWN"}

EVIDENCE_PROPOSITION_TYPES = {
    "JURISDICTION", "AUTHORITY_HIERARCHY", "CODE_CURRENCY", "APPLICABILITY",
    "PERMIT_REQUIREMENT", "REVIEW_REQUIREMENT", "PERMIT_EXEMPTION",
    "PATHWAY", "THRESHOLD", "ENTITLEMENT", "OTHER",
}

GEMINI_KEY = os.getenv("GEMINI_KEY") or st.secrets.get("GEMINI_KEY", "")
PROMPT_VERSION = "v-test-harness-1.0"

# ============================================================
# CORE HELPERS (deterministic contract)
# ============================================================
def _norm_text(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())

def discipline_family(value):
    value = str(value or "").strip().lower()
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
        "plumbing": ["plumbing", "plumbing"],
    }
    for family, terms in families.items():
        if any(term in value for term in terms):
            return family
    return value

def evidence_proposition_integrity_errors(evidence):
    errors = []
    if not evidence:
        return errors
    eid = evidence.get("id", "Unknown")
    ptype = evidence.get("proposition_type")
    rule = _norm_text(evidence.get("rule"))
    if not rule:
        return errors

    permit_patterns = [
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

    if ptype == "PERMIT_REQUIREMENT" and not any(p.search(rule) for p in permit_patterns):
        errors.append(f"{eid}: PERMIT_REQUIREMENT label unsupported by rule text.")
    if ptype == "PERMIT_EXEMPTION" and not exemption_explicit.search(rule):
        errors.append(f"{eid}: PERMIT_EXEMPTION label unsupported by rule text.")
    if ptype == "REVIEW_REQUIREMENT" and not review_terms.search(rule):
        errors.append(f"{eid}: REVIEW_REQUIREMENT label unsupported.")
    if ptype == "PATHWAY" and not pathway_terms.search(rule):
        errors.append(f"{eid}: PATHWAY label unsupported.")
    if ptype == "THRESHOLD" and not threshold_terms.search(rule):
        errors.append(f"{eid}: THRESHOLD label unsupported.")
    return errors

def evidence_ids_supporting_type(evidence_ids, evidence_by_id, proposition_type, discipline):
    result = []
    for eid in evidence_ids or []:
        evidence = evidence_by_id.get(eid)
        if not evidence:
            continue
        if discipline_family(evidence.get("discipline", "")) != discipline_family(discipline):
            continue
        if evidence.get("proposition_type") != proposition_type:
            continue
        if evidence_proposition_integrity_errors(evidence):
            continue
        result.append(eid)
    return result

def derive_permit_statuses_from_evidence(data):
    """Single source of truth for permit status."""
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
        valid_req = evidence_ids_supporting_type(ids, evidence_by_id, "PERMIT_REQUIREMENT", discipline)
        valid_exempt = evidence_ids_supporting_type(ids, evidence_by_id, "PERMIT_EXEMPTION", discipline)

        if valid_exempt:
            item["permit"] = "NOT_APPLICABLE"
            item["permit_basis"] = "DIRECT_EVIDENCE"
            item["permit_evidence"] = valid_exempt
            item["permit_finding"] = (
                f"Authoritative evidence establishes that a {discipline.lower()} "
                "permit is not required for the stated scope."
            )
        elif valid_req:
            item["permit"] = "VERIFIED_REQUIRED"
            item["permit_basis"] = "DIRECT_EVIDENCE"
            item["permit_evidence"] = valid_req
            item["permit_finding"] = (
                f"Authoritative evidence establishes a {discipline.lower()} "
                "permit requirement for the stated scope."
            )
        else:
            current = str(item.get("permit") or "UNKNOWN").upper()
            if current in {"VERIFIED_REQUIRED", "REQUIRED", "INFERRED", "NOT_APPLICABLE"}:
                item["permit"] = "CONDITIONAL"
            item["permit_basis"] = "NOT_ESTABLISHED"
            item["permit_evidence"] = []
            item["permit_finding"] = (
                f"Current evidence does not yet establish whether a {discipline.lower()} "
                "permit is required. Confirm with an authoritative permit-specific source."
            )
    return data

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
    ]
    # Very light site-specificity check for the harness
    site_specific = False
    for e in valid:
        text = _norm_text(" ".join([
            str(e.get("title") or ""),
            str(e.get("rule") or ""),
            str(e.get("retrieval_note") or ""),
        ]))
        if "unincorporated" in text or "city limits" in text or "jurisdiction" in text:
            site_specific = True
            break
    if not site_specific:
        jurisdiction["status"] = "CONDITIONAL"
        jurisdiction["validation_note"] = (
            "Jurisdiction downgraded: cited evidence did not establish "
            "parcel-level governmental boundary result."
        )
    return data

def run_deterministic_contract_pass(data, address_text="", project_date_value=None, selected_state=""):
    """Minimal but real contract pass for testing."""
    if not isinstance(data, dict):
        return data, ["Dossier is not a JSON object."]
    data = sanitize_unverifiable_verified_jurisdiction(data, str(address_text or ""))
    data = derive_permit_statuses_from_evidence(data)
    # Light validation
    errors = []
    for item in data.get("disciplines") or []:
        permit = str(item.get("permit") or "").upper()
        if permit == "VERIFIED_REQUIRED":
            ids = item.get("permit_evidence") or []
            if not ids:
                errors.append(f"{item.get('type')}: VERIFIED_REQUIRED without permit_evidence.")
    return data, errors

# ============================================================
# MOCK RESEARCH RESULTS
# ============================================================
def build_mock_dossier_la_puente():
    """Realistic mock for the large La Puente-style SOW."""
    return {
        "bottom_line": (
            "Jurisdiction remains conditional pending parcel-level confirmation "
            "(La Puente postal city vs possible unincorporated Los Angeles County). "
            "Multiple disciplines (Building, Electrical, Mechanical, Plumbing, Accessibility, "
            "Fire) have clear work scopes that typically trigger permits, but proposition-specific "
            "permit evidence has not yet been locked. Accessibility alterations and existing "
            "entitlement conditions are material open items."
        ),
        "bottom_line_evidence": ["E1", "E2"],
        "research_completeness": {
            "status": "PARTIAL",
            "reason": "Main framework visible; jurisdiction boundary, existing CUP/entitlements, and several permit-specific sources still needed.",
            "critical_missing": [
                "Parcel-level jurisdiction / municipal boundary confirmation",
                "Existing CUP or site-plan conditions (if any)",
                "Equipment cut sheets (HVAC weight, electrical MCA/MOP)",
                "Accessibility alteration level / path-of-travel scope",
            ],
        },
        "jurisdiction": {
            "status": "CONDITIONAL",
            "county": "Los Angeles County",
            "city": "La Puente (postal city – boundary unconfirmed)",
            "ahj": "Unresolved – confirm City of La Puente vs unincorporated LA County",
            "evidence": ["E1"],
            "validation_note": "Postal city alone does not establish municipal jurisdiction.",
        },
        "codes": [
            {"name": "2022 California Building Code (CBC)", "status": "CONDITIONAL", "evidence": ["E3"]},
            {"name": "2022 California Electrical Code", "status": "CONDITIONAL", "evidence": ["E3"]},
            {"name": "2022 California Mechanical Code", "status": "CONDITIONAL", "evidence": ["E3"]},
            {"name": "2022 California Plumbing Code", "status": "CONDITIONAL", "evidence": ["E3"]},
        ],
        "evidence": [
            {
                "id": "E1",
                "title": "Los Angeles County / City boundary research note",
                "url": "https://www.lacounty.gov",
                "authority": "county",
                "discipline": "Jurisdiction",
                "proposition_type": "JURISDICTION",
                "source_type": "other",
                "retrieval_note": "Postal city is La Puente; many nearby parcels are unincorporated.",
                "rule": "Governmental jurisdiction requires parcel-level confirmation of municipal vs unincorporated status.",
            },
            {
                "id": "E2",
                "title": "Typical California commercial remodel permit triggers",
                "url": "https://www.dgs.ca.gov/BSC",
                "authority": "state",
                "discipline": "Building",
                "proposition_type": "APPLICABILITY",
                "source_type": "code",
                "retrieval_note": "General applicability only – not a permit requirement.",
                "rule": "Alterations to existing buildings are subject to the California Building Standards Code.",
            },
            {
                "id": "E3",
                "title": "California Building Standards Commission – current codes",
                "url": "https://www.dgs.ca.gov/BSC",
                "authority": "state",
                "discipline": "Building",
                "proposition_type": "CODE_CURRENCY",
                "source_type": "code",
                "retrieval_note": "State code adoption reference.",
                "rule": "2022 California codes are the current statewide baseline pending local amendments.",
            },
            {
                "id": "E4",
                "title": "Electrical work – panel and branch circuit replacement",
                "url": "https://www.dgs.ca.gov/BSC",
                "authority": "state",
                "discipline": "Electrical",
                "proposition_type": "APPLICABILITY",
                "source_type": "code",
                "retrieval_note": "Applicability only.",
                "rule": "Replacement of electrical panels and branch circuits is subject to the California Electrical Code.",
            },
            {
                "id": "E5",
                "title": "Mechanical equipment replacement",
                "url": "https://www.dgs.ca.gov/BSC",
                "authority": "state",
                "discipline": "Mechanical",
                "proposition_type": "APPLICABILITY",
                "source_type": "code",
                "retrieval_note": "Applicability only.",
                "rule": "Replacement of HVAC equipment is subject to the California Mechanical Code and energy provisions.",
            },
        ],
        "disciplines": [
            {
                "type": "Building / Structural",
                "applicability": {
                    "rule": "Alterations and repairs to existing buildings are subject to the CBC.",
                    "fact": {"statement": "Exterior repairs, window infill, roof replacement, interior wall/platform adjustments.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "",
                    "relationship": "direct",
                    "evidence": ["E2"],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Building permit is typically required for this scope of alteration, but proposition-specific local permit evidence is not yet established.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Plan-check vs over-the-counter pathway not yet established for this AHJ.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Exact structural repair extent", "Roof structure vs cover-only determination"],
                "reopen": ["Structural damage discovered beyond repair-in-kind"],
                "actionable_questions": [
                    "Does the AHJ treat window infill + stucco repair + full roof cover replacement as a single building permit or multiple permits?",
                    "Is a structural engineer’s letter required when roof structure repair is ‘only if required by code’?",
                ],
                "potential_issues": [
                    {"issue": "Roof structure repair may trigger structural plan review if more than sheathing is involved.", "why": "SOW language leaves structural scope open."},
                ],
            },
            {
                "type": "Electrical",
                "applicability": {
                    "rule": "Panel, branch circuit, and fixture replacement is subject to the CEC.",
                    "fact": {"statement": "Full replacement of meter/panel, branch circuits, switches, receptacles, and fixtures.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "",
                    "relationship": "direct",
                    "evidence": ["E4"],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Full electrical system replacement almost always requires an electrical permit; local permit-specific evidence still needed.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Separate electrical permit vs combined building permit pathway not yet confirmed.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Service size / panel ampacity", "Whether service lateral is also being replaced"],
                "reopen": [],
                "actionable_questions": [
                    "Will the new panel be the same ampacity and location, or is a service upgrade involved?",
                    "Does the AHJ require a separate electrical permit for full branch-circuit replacement when a building permit is also pulled?",
                ],
                "potential_issues": [],
            },
            {
                "type": "Mechanical",
                "applicability": {
                    "rule": "HVAC replacement is subject to the CMC and energy code.",
                    "fact": {"statement": "Replace HVAC with ground-mounted system if possible; replace ductwork.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "Final equipment location, weight, and efficiency ratings.",
                    "relationship": "direct",
                    "evidence": ["E5"],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Mechanical permit is expected for equipment and ductwork replacement; ground-mount location may add structural and planning considerations.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway depends on whether units remain roof-mounted or move to grade.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Equipment weight and anchorage", "Noise / setback implications of ground-mount"],
                "reopen": ["Ground-mount location changes site or parking layout"],
                "actionable_questions": [
                    "If units move from roof to grade, does the AHJ require structural calculations for the new pads and any screening?",
                    "Does an existing CUP or site plan condition restrict exterior mechanical equipment location or noise?",
                ],
                "potential_issues": [
                    {"issue": "Ground-mounted units may trigger planning or noise review depending on setbacks and existing entitlements.", "why": "SOW prefers ground-mount if possible."},
                ],
            },
            {
                "type": "Plumbing",
                "applicability": {
                    "rule": "Fixture and water-heater replacement plus accessibility-driven sanitary adjustments are subject to the CPC.",
                    "fact": {"statement": "Replace toilets, sinks, water heaters, fountains; adjust sanitary drainage for accessibility.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "",
                    "relationship": "direct",
                    "evidence": [],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Plumbing permit expected for fixture and water-heater replacement and drainage modifications.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Whether water service or sewer lateral work extends beyond the building"],
                "reopen": [],
                "actionable_questions": [
                    "Is the sewer lateral replacement (pending video) inside or outside the building footprint, and which AHJ reviews it?",
                ],
                "potential_issues": [],
            },
            {
                "type": "Accessibility",
                "applicability": {
                    "rule": "Alterations that affect accessibility features must comply with current CBC Chapter 11B requirements.",
                    "fact": {"statement": "Lobby, bathroom, platform, railings, sidewalks, and ADA parking are being adjusted.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "Exact alteration level and whether path-of-travel obligations are triggered.",
                    "relationship": "direct",
                    "evidence": [],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Accessibility compliance is required; whether a separate accessibility review or inclusion in the building permit is not yet established.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Path-of-travel evaluation", "Primary function area determination"],
                "reopen": [],
                "actionable_questions": [
                    "Does the AHJ require a separate accessibility plan review or is it handled inside the building permit for this scope?",
                    "Will the lobby and bathroom work be treated as an alteration to a primary function area that triggers additional path-of-travel upgrades?",
                ],
                "potential_issues": [
                    {"issue": "Path-of-travel obligations may extend beyond the rooms being directly altered.", "why": "SOW includes lobby, bathrooms, platform, sidewalks, and parking."},
                ],
            },
            {
                "type": "Fire / Life Safety",
                "applicability": {
                    "rule": "Smoke/CO detector replacement and any fire-alarm work are subject to fire-code provisions.",
                    "fact": {"statement": "Replace smoke/CO detectors and security alarm system throughout.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "Whether the security system is interconnected with fire alarm.",
                    "relationship": "direct",
                    "evidence": [],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Detector replacement often requires a fire permit or notification; exact local requirement not yet locked.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Fire-alarm vs detection-only scope"],
                "reopen": [],
                "actionable_questions": [
                    "Does detector-only replacement require a fire permit or can it be performed under the building permit with inspection?",
                ],
                "potential_issues": [],
            },
            {
                "type": "Planning / Land Use",
                "applicability": {
                    "rule": "Exterior changes, parking re-striping, and mechanical location may be subject to existing entitlements or site-plan conditions.",
                    "fact": {"statement": "Exterior work, parking lot re-stripe with ADA stalls, possible ground-mounted HVAC, fencing/CMU repairs.", "source": "USER_PROVIDED"},
                    "determination": "cannot_determine",
                    "missing": "Whether a CUP, site plan, or other entitlement governs the site.",
                    "relationship": "conditional",
                    "evidence": ["E1"],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Planning clearance or CUP consistency review may be required; existing entitlement documents have not been retrieved.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway depends on existence and conditions of any current entitlement.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Existing CUP / site-plan / development agreement documents"],
                "reopen": ["Ground-mount HVAC or parking changes conflict with recorded conditions"],
                "actionable_questions": [
                    "Does the property have an active CUP, site plan, or other land-use approval that governs exterior equipment, parking, or signage?",
                    "Is a planning clearance required before building permit issuance for the parking re-stripe and any new mechanical pads?",
                ],
                "potential_issues": [
                    {"issue": "Existing entitlement conditions may restrict exterior mechanical placement or require design review.", "why": "SOW contemplates possible ground-mounted units and parking changes."},
                ],
            },
        ],
    }


def build_mock_dossier_small_furnace():
    """Simple mock for a small residential furnace replacement."""
    return {
        "bottom_line": (
            "Mechanical permit is typically required for furnace replacement. "
            "Electrical scope (disconnect / circuit) and any combustion-air or flue changes "
            "remain to be confirmed. Jurisdiction is treated as the selected small-town AHJ."
        ),
        "bottom_line_evidence": ["E1"],
        "research_completeness": {
            "status": "PARTIAL",
            "reason": "Core mechanical trigger visible; local permit form and any electrical work still open.",
            "critical_missing": ["Exact electrical scope", "Local mechanical permit application"],
        },
        "jurisdiction": {
            "status": "CONDITIONAL",
            "county": "Example County",
            "city": "Small Town",
            "ahj": "Small Town Building Department",
            "evidence": [],
        },
        "codes": [
            {"name": "Current state mechanical code", "status": "CONDITIONAL", "evidence": []},
        ],
        "evidence": [
            {
                "id": "E1",
                "title": "Typical mechanical permit trigger for furnace replacement",
                "url": "https://example.com",
                "authority": "state",
                "discipline": "Mechanical",
                "proposition_type": "APPLICABILITY",
                "source_type": "code",
                "retrieval_note": "General applicability.",
                "rule": "Replacement of a fuel-fired furnace is subject to the mechanical code.",
            },
        ],
        "disciplines": [
            {
                "type": "Mechanical",
                "applicability": {
                    "rule": "Furnace replacement is subject to the mechanical code.",
                    "fact": {"statement": "Replace existing gas furnace.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "missing": "",
                    "relationship": "direct",
                    "evidence": ["E1"],
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Mechanical permit is expected; local permit-specific evidence not yet retrieved.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Whether flue or combustion air is being altered"],
                "reopen": [],
                "actionable_questions": [
                    "Does the local AHJ allow over-the-counter mechanical permits for like-for-like furnace replacement?",
                ],
                "potential_issues": [],
            },
            {
                "type": "Electrical",
                "applicability": {
                    "rule": "Electrical work associated with furnace replacement may require a permit.",
                    "fact": {"statement": "Electrical scope not stated.", "source": "USER_PROVIDED"},
                    "determination": "cannot_determine",
                    "missing": "Whether disconnect, circuit, or wiring is being modified.",
                    "relationship": "conditional",
                    "evidence": [],
                },
                "permit": "UNKNOWN",
                "permit_finding": "Electrical permit consequence cannot be determined from current facts.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "UNKNOWN",
                "pathway_finding": "Pathway cannot be established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Electrical scope"],
                "reopen": [],
                "actionable_questions": [
                    "Will the furnace replacement include a new disconnect, circuit, or wiring changes?",
                ],
                "potential_issues": [],
            },
        ],
    }


# ============================================================
# UI
# ============================================================
if "report_data" not in st.session_state:
    st.session_state.report_data = None
if "debug_log" not in st.session_state:
    st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state:
    st.session_state.error_msg = None
if "run_status" not in st.session_state:
    st.session_state.run_status = "Ready"

st.title("🏛️ AHJ Research Assistant – Test Harness")
st.caption("Deterministic evidence contract + mock research path. Safe to run without a Gemini key.")

with st.sidebar:
    st.warning("Mock Mode is recommended for testing. Live Gemini calls cost money.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=True)
    st.markdown("---")
    st.markdown("**Quick test presets**")
    if st.button("Load La Puente large SOW"):
        st.session_state["preset"] = "la_puente"
    if st.button("Load small furnace SOW"):
        st.session_state["preset"] = "furnace"

# Defaults
default_state = "California"
default_address = "13431 Temple Ave, La Puente, CA"
default_sow = """EXTERIOR BUILDING – Structural Frame repair per discovery; Exterior Wall Construction repair + window infill; Exterior Siding/stucco repair; Exterior Paint; Roof Cover replace; HVAC Units replace with ground-mounted system if possible; Electrical – full panel, branch circuits, fixtures replace; Plumbing – toilets, sinks, water heaters replace; Accessibility adjustments to lobby, bathrooms, platform, sidewalks, ADA parking; etc. (full multi-discipline remodel)."""

preset = st.session_state.get("preset")
if preset == "furnace":
    default_state = "Montana"
    default_address = "123 Main St, Small Town, MT"
    default_sow = "Replace existing gas furnace with like-for-like unit. No ductwork changes. Electrical scope unknown."

st.header("1. Project Metadata")
col1, col2 = st.columns(2)
with col1:
    state = st.selectbox("State / Jurisdiction", STATE_OPTIONS, index=STATE_OPTIONS.index(default_state))
    address = st.text_input("Project Address", default_address)
    project_date = st.date_input("Permit / Construction Date", date.today())
with col2:
    ptype = st.selectbox("Project Type", PROJECT_TYPES, index=1)
    bclass = st.selectbox("Building / Occupancy Class", BUILDING_CLASSES, index=1)
    existing_permit = st.text_input("Existing Entitlements (Optional)", "", placeholder="e.g., CUP, variance")

st.header("2. Scope of Work (SOW)")
sow_text = st.text_area("Paste the complete Scope of Work below.", height=220, value=default_sow)

st.header("3. Research Execution")
if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    st.session_state.error_msg = None
    st.session_state.report_data = None
    st.session_state.run_status = "Started"

    if mock_mode:
        with st.spinner("Running mock research + deterministic contract..."):
            if "la puente" in address.lower() or "temple ave" in address.lower():
                raw = build_mock_dossier_la_puente()
            else:
                raw = build_mock_dossier_small_furnace()

            data, errors = run_deterministic_contract_pass(raw, address, project_date, state)
            st.session_state.report_data = data
            st.session_state.debug_log = {
                "mock": True,
                "validation_errors": errors,
                "note": "Mock data passed through deterministic contract.",
            }
            st.session_state.run_status = "Complete – mock dossier ready"
            if errors:
                st.session_state.error_msg = "Deterministic contract reported issues (see debug)."
    else:
        if not GEMINI_KEY or not GENAI_AVAILABLE:
            st.session_state.error_msg = "GEMINI_KEY missing or google-genai not installed. Use Mock Mode."
            st.session_state.run_status = "Blocked"
        else:
            st.session_state.error_msg = (
                "Live Gemini path is intentionally minimal in this test harness. "
                "Use your full production app.py for real research, or stay in Mock Mode."
            )
            st.session_state.run_status = "Live path not fully wired in harness"

st.info(f"Research status: {st.session_state.run_status}")

if st.session_state.error_msg:
    st.error(st.session_state.error_msg)

with st.expander("🐛 Debug Log", expanded=bool(st.session_state.error_msg)):
    st.json(st.session_state.debug_log)

# ============================================================
# RESULTS
# ============================================================
if st.session_state.report_data:
    data = st.session_state.report_data
    st.divider()
    st.header("4. Research Dossier")

    st.info(f"**Bottom Line:** {data.get('bottom_line', 'N/A')}")
    completeness = data.get("research_completeness") or {}
    if completeness:
        status = completeness.get("status", "UNKNOWN")
        if status == "PARTIAL":
            st.warning(f"Research completeness: PARTIAL — {completeness.get('reason', '')}")
        elif status == "SUFFICIENT":
            st.success("Research completeness: SUFFICIENT")
        else:
            st.error(f"Research completeness: {status}")

    st.subheader("📍 Jurisdiction")
    jur = data.get("jurisdiction") or {}
    st.write(f"**Status:** {jur.get('status', 'N/A')}")
    st.write(f"**County:** {jur.get('county', 'Unknown')}")
    st.write(f"**City:** {jur.get('city', 'Unknown')}")
    st.write(f"**AHJ:** {jur.get('ahj', 'Unknown')}")
    if jur.get("validation_note"):
        st.caption(jur["validation_note"])

    st.subheader("📋 Discipline Summary")
    for item in data.get("disciplines") or []:
        permit = str(item.get("permit", "UNKNOWN")).upper()
        icon = {"VERIFIED_REQUIRED": "🟢", "CONDITIONAL": "🟠", "UNKNOWN": "🟡", "NOT_APPLICABLE": "⚪"}.get(permit, "🟡")
        with st.expander(f"{icon} {item.get('type', 'Unknown')} — {permit}"):
            st.markdown(f"**Permit finding:** {item.get('permit_finding', 'N/A')}")
            st.markdown(f"**Pathway finding:** {item.get('pathway_finding', 'N/A')}")
            questions = item.get("actionable_questions") or []
            if questions:
                st.markdown("**Questions to resolve next:**")
                for q in questions:
                    st.markdown(f"- {q}")
            leads = item.get("potential_issues") or []
            if leads:
                st.markdown("**Worth checking (AI lead):**")
                for lead in leads:
                    if isinstance(lead, dict):
                        st.markdown(f"- {lead.get('issue', '')}")
                        if lead.get("why"):
                            st.caption(f"Why: {lead['why']}")

    st.header("5. Export")
    json_data = json.dumps(
        {"project": {"state": state, "address": address, "date": str(project_date)}, "dossier": data},
        indent=2,
    )
    st.download_button(
        "💾 Download JSON",
        data=json_data,
        file_name="AHJ_Test_Dossier.json",
        mime="application/json",
        use_container_width=True,
    )
