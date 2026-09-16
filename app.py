# app.py
# AHJ Research Assistant – Clean Production Version
# Incorporates: evidence taxonomy, deterministic permit engine,
# jurisdiction firewall, source-specificity checks, mock + live Gemini paths

import os
import re
import json
import time
import hashlib
import logging
from datetime import datetime, date, timezone
from io import BytesIO

import streamlit as st

# Optional dependencies
try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

try:
    from docx import Document
    from docx.shared import Pt, Inches
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

st.set_page_config(
    page_title="AHJ Research Assistant",
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

EVIDENCE_PROPOSITION_TYPES = {
    "JURISDICTION", "AUTHORITY_HIERARCHY", "CODE_CURRENCY", "APPLICABILITY",
    "PERMIT_REQUIREMENT", "PERMIT_EXEMPTION", "REVIEW_REQUIREMENT",
    "PATHWAY", "THRESHOLD", "ENTITLEMENT", "OTHER",
}

FACT_SOURCES = {"USER_PROVIDED", "RETRIEVED_RECORD", "AUTHORITATIVE_SOURCE", "INFERRED", "UNKNOWN"}

GEMINI_KEY = os.getenv("GEMINI_KEY") or st.secrets.get("GEMINI_KEY", "")
PROMPT_VERSION = "v27.0-clean"

# ============================================================
# TEXT / DISCIPLINE HELPERS
# ============================================================
def _norm(value):
    return re.sub(r"\s+", " ", str(value or "").strip().lower())

def discipline_family(value):
    value = _norm(value)
    if "structural" in value:
        return "structural"
    if "accessibility" in value or "ada" in value:
        return "accessibility"
    families = {
        "mechanical": ["mechanical", "hvac", "heating", "cooling"],
        "electrical": ["electrical", "electric"],
        "plumbing": ["plumbing"],
        "energy": ["energy"],
        "planning": ["planning", "land use", "zoning", "cup"],
        "fire": ["fire", "life safety"],
        "building": ["building", "construction"],
    }
    for family, terms in families.items():
        if any(t in value for t in terms):
            return family
    return value

# ============================================================
# EVIDENCE CONTRACT (core of the system)
# ============================================================
def evidence_proposition_integrity_errors(ev):
    """Reject evidence whose label is stronger than its own rule text."""
    if not isinstance(ev, dict):
        return ["Invalid evidence object"]
    errors = []
    eid = ev.get("id", "Evidence")
    ptype = str(ev.get("proposition_type") or "").upper()
    rule = _norm(ev.get("rule"))

    if not rule:
        return errors

    permit_re = re.compile(
        r"\b(?:permit|approval|license)\b[^.!?;:]{0,180}\b(?:required|needed|necessary)\b|"
        r"\b(?:requires?|must obtain|shall obtain)\b[^.!?;:]{0,180}\b(?:permit|approval|license)\b",
        re.I,
    )
    exempt_re = re.compile(
        r"\bexempt(?:ed|ion)?\b|\bno\s+(?:[a-z-]+\s+)?permit\b|"
        r"\bpermit\s+(?:is\s+)?not\s+required\b|\bdoes\s+not\s+require\s+(?:a\s+)?permit\b",
        re.I,
    )
    review_re = re.compile(r"\breview\b|\binspection\b|\bsubmittal\b|\bplan review\b", re.I)
    pathway_re = re.compile(
        r"\bsubmit\b|\bapplication\b|\bportal\b|\bonline\b|\bover[- ]the[- ]counter\b|"
        r"\bprocess(?:ed|ing)?\b|\bfil(?:e|ing)\b|\bpermit center\b|\bplan review\b",
        re.I,
    )

    if ptype == "PERMIT_REQUIREMENT" and not permit_re.search(rule):
        errors.append(f"{eid}: PERMIT_REQUIREMENT label not supported by rule text.")
    if ptype == "PERMIT_EXEMPTION" and not exempt_re.search(rule):
        errors.append(f"{eid}: PERMIT_EXEMPTION label not supported by rule text.")
    if ptype == "REVIEW_REQUIREMENT" and not review_re.search(rule):
        errors.append(f"{eid}: REVIEW_REQUIREMENT label not supported by rule text.")
    if ptype == "PATHWAY" and not pathway_re.search(rule):
        errors.append(f"{eid}: PATHWAY label not supported by rule text.")
    return errors


def evidence_ids_supporting_type(ids, evidence_by_id, prop_type, discipline):
    """Return only evidence that is valid for the requested proposition + discipline."""
    good = []
    for eid in ids or []:
        ev = evidence_by_id.get(str(eid))
        if not ev:
            continue
        if discipline_family(ev.get("discipline", "")) != discipline_family(discipline):
            continue
        if str(ev.get("proposition_type") or "").upper() != prop_type:
            continue
        if evidence_proposition_integrity_errors(ev):
            continue
        good.append(str(eid))
    return good


def derive_permit_statuses_from_evidence(data):
    """
    Single source of truth for permit status.
    Gemini may discover evidence; only this function decides the final status.
    """
    if not isinstance(data, dict):
        return data

    evidence_by_id = {
        str(e.get("id")): e
        for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }

    for item in data.get("disciplines") or []:
        if not isinstance(item, dict):
            continue
        discipline = str(item.get("type") or "Unknown")
        ids = item.get("permit_evidence") or []
        if isinstance(ids, str):
            ids = [ids]

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


def sanitize_jurisdiction(data, address=""):
    """Downgrade VERIFIED jurisdiction that lacks site-specific evidence."""
    if not isinstance(data, dict):
        return data
    jur = data.get("jurisdiction") or {}
    if str(jur.get("status") or "").upper() != "VERIFIED":
        return data

    evidence_by_id = {
        str(e.get("id")): e
        for e in (data.get("evidence") or [])
        if isinstance(e, dict) and e.get("id")
    }
    ids = jur.get("evidence") or []
    site_specific = False
    for eid in ids:
        ev = evidence_by_id.get(str(eid))
        if not ev or str(ev.get("proposition_type") or "").upper() != "JURISDICTION":
            continue
        text = _norm(
            " ".join([
                str(ev.get("title") or ""),
                str(ev.get("rule") or ""),
                str(ev.get("retrieval_note") or ""),
            ])
        )
        if any(t in text for t in [
            "unincorporated", "city limits", "municipal limits",
            "jurisdiction is", "within the city", "outside the city",
            "county jurisdiction", "parcel",
        ]):
            site_specific = True
            break

    if not site_specific:
        jur["status"] = "CONDITIONAL"
        jur["validation_note"] = (
            "Jurisdiction downgraded: cited evidence did not establish "
            "parcel-level governmental boundary result. Postal city alone is insufficient."
        )
        # Special handling for La Puente-style cases
        if "la puente" in _norm(address):
            jur["city"] = "La Puente (postal – boundary unconfirmed)"
            jur["ahj"] = "Confirm City of La Puente vs unincorporated Los Angeles County"
    data["jurisdiction"] = jur
    return data


def run_deterministic_contract(data, address="", state=""):
    """Final deterministic pass before the dossier is shown to the user."""
    if not isinstance(data, dict):
        return data, ["Dossier is not a JSON object."]

    data = sanitize_jurisdiction(data, address)
    data = derive_permit_statuses_from_evidence(data)

    errors = []
    for item in data.get("disciplines") or []:
        if str(item.get("permit") or "").upper() == "VERIFIED_REQUIRED":
            if not item.get("permit_evidence"):
                errors.append(
                    f"{item.get('type')}: VERIFIED_REQUIRED without valid permit_evidence."
                )
    return data, errors


# ============================================================
# MOCK RESEARCH (safe, free)
# ============================================================
def mock_dossier_large(address):
    """Realistic multi-discipline mock for large remodel / La Puente-style jobs."""
    is_la_puente = "la puente" in _norm(address) or "temple ave" in _norm(address)
    return {
        "bottom_line": (
            "Jurisdiction remains conditional pending parcel-level confirmation. "
            "Multiple disciplines have clear work that typically triggers permits, "
            "but proposition-specific local permit evidence is not yet locked. "
            "Accessibility alterations and any existing entitlements are material open items."
        ),
        "bottom_line_evidence": ["E1", "E2"],
        "research_completeness": {
            "status": "PARTIAL",
            "reason": "Framework visible; jurisdiction boundary, existing entitlements, and several permit-specific sources still needed.",
            "critical_missing": [
                "Parcel-level jurisdiction confirmation",
                "Existing CUP / site-plan conditions (if any)",
                "Equipment cut sheets (weight, MCA/MOP)",
                "Accessibility path-of-travel scope",
            ],
        },
        "jurisdiction": {
            "status": "CONDITIONAL",
            "county": "Los Angeles County" if is_la_puente else "Example County",
            "city": "La Puente (postal – boundary unconfirmed)" if is_la_puente else "Example City",
            "ahj": "Confirm City vs unincorporated County" if is_la_puente else "Local Building Department",
            "evidence": ["E1"],
        },
        "codes": [
            {"name": "2022 California Building Code", "status": "CONDITIONAL", "evidence": ["E3"]},
            {"name": "2022 California Electrical Code", "status": "CONDITIONAL", "evidence": ["E3"]},
            {"name": "2022 California Mechanical Code", "status": "CONDITIONAL", "evidence": ["E3"]},
        ],
        "evidence": [
            {
                "id": "E1",
                "title": "Jurisdiction boundary research note",
                "url": "https://www.lacounty.gov",
                "authority": "county",
                "discipline": "Jurisdiction",
                "proposition_type": "JURISDICTION",
                "source_type": "other",
                "retrieval_note": "Postal city alone does not establish municipal jurisdiction.",
                "rule": "Governmental jurisdiction requires parcel-level confirmation of municipal versus unincorporated status.",
            },
            {
                "id": "E2",
                "title": "California commercial alteration applicability",
                "url": "https://www.dgs.ca.gov/BSC",
                "authority": "state",
                "discipline": "Building",
                "proposition_type": "APPLICABILITY",
                "source_type": "code",
                "retrieval_note": "Applicability only – not a permit requirement.",
                "rule": "Alterations to existing buildings are subject to the California Building Standards Code.",
            },
            {
                "id": "E3",
                "title": "California Building Standards Commission – code adoption",
                "url": "https://www.dgs.ca.gov/BSC",
                "authority": "state",
                "discipline": "Building",
                "proposition_type": "CODE_CURRENCY",
                "source_type": "code",
                "retrieval_note": "State baseline reference.",
                "rule": "2022 California codes are the current statewide baseline pending local amendments.",
            },
        ],
        "disciplines": [
            {
                "type": "Building / Structural",
                "applicability": {
                    "rule": "Alterations and repairs are subject to the CBC.",
                    "fact": {
                        "statement": "Exterior repairs, window infill, roof replacement, interior adjustments.",
                        "source": "USER_PROVIDED",
                    },
                    "determination": "applies",
                    "relationship": "direct",
                    "evidence": ["E2"],
                    "missing": "",
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Building permit typically required; local permit-specific evidence not yet established.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Plan-check pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Exact structural repair extent"],
                "reopen": ["Damage beyond repair-in-kind discovered"],
                "actionable_questions": [
                    "Does the AHJ treat window infill + roof cover replacement as one building permit or multiple?",
                    "Is a structural letter required when roof structure repair is left open-ended?",
                ],
                "potential_issues": [
                    {
                        "issue": "Roof structure work may trigger structural plan review.",
                        "why": "SOW leaves structural scope open-ended.",
                    }
                ],
            },
            {
                "type": "Electrical",
                "applicability": {
                    "rule": "Panel and branch-circuit replacement is subject to the CEC.",
                    "fact": {
                        "statement": "Full replacement of panel, branch circuits, and fixtures.",
                        "source": "USER_PROVIDED",
                    },
                    "determination": "applies",
                    "relationship": "direct",
                    "evidence": [],
                    "missing": "",
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Full electrical replacement almost always requires a permit; local evidence still needed.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Separate electrical vs combined building permit not yet confirmed.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Service size / whether service lateral is included"],
                "reopen": [],
                "actionable_questions": [
                    "Will the new panel keep the same ampacity and location, or is a service upgrade involved?",
                ],
                "potential_issues": [],
            },
            {
                "type": "Mechanical",
                "applicability": {
                    "rule": "HVAC replacement is subject to the CMC and energy provisions.",
                    "fact": {
                        "statement": "Replace HVAC; prefer ground-mounted if possible; replace ductwork.",
                        "source": "USER_PROVIDED",
                    },
                    "determination": "applies",
                    "relationship": "direct",
                    "evidence": [],
                    "missing": "Final location, weight, efficiency ratings.",
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Mechanical permit expected; ground-mount may add structural/planning issues.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway depends on roof vs grade location.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Equipment weight and anchorage details"],
                "reopen": ["Ground-mount changes site layout"],
                "actionable_questions": [
                    "If units move to grade, does the AHJ require structural calculations and any screening?",
                    "Do existing entitlements restrict exterior mechanical location or noise?",
                ],
                "potential_issues": [
                    {
                        "issue": "Ground-mounted units may trigger planning or noise review.",
                        "why": "SOW prefers ground-mount if possible.",
                    }
                ],
            },
            {
                "type": "Accessibility",
                "applicability": {
                    "rule": "Alterations affecting accessibility must comply with CBC Chapter 11B.",
                    "fact": {
                        "statement": "Lobby, bathrooms, platform, railings, sidewalks, and ADA parking are being adjusted.",
                        "source": "USER_PROVIDED",
                    },
                    "determination": "applies",
                    "relationship": "direct",
                    "evidence": [],
                    "missing": "Exact alteration level and path-of-travel obligations.",
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Accessibility compliance required; separate review vs inclusion in building permit not yet established.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Path-of-travel evaluation"],
                "reopen": [],
                "actionable_questions": [
                    "Does the AHJ require a separate accessibility plan review for this scope?",
                    "Will lobby/bathroom work trigger additional path-of-travel upgrades?",
                ],
                "potential_issues": [
                    {
                        "issue": "Path-of-travel obligations may extend beyond the rooms directly altered.",
                        "why": "Multiple accessibility-related scopes are present.",
                    }
                ],
            },
            {
                "type": "Planning / Land Use",
                "applicability": {
                    "rule": "Exterior changes and parking work may be subject to existing entitlements.",
                    "fact": {
                        "statement": "Exterior work, parking re-stripe, possible ground-mounted HVAC.",
                        "source": "USER_PROVIDED",
                    },
                    "determination": "cannot_determine",
                    "relationship": "conditional",
                    "evidence": ["E1"],
                    "missing": "Whether a CUP or site plan governs the property.",
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Planning clearance or CUP consistency review may be required; governing documents not retrieved.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Depends on existence of current entitlements.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Existing CUP / site-plan documents"],
                "reopen": ["Ground-mount or parking changes conflict with recorded conditions"],
                "actionable_questions": [
                    "Does the property have an active CUP or site-plan approval?",
                    "Is planning clearance required before building permit issuance?",
                ],
                "potential_issues": [
                    {
                        "issue": "Existing entitlement conditions may restrict exterior equipment placement.",
                        "why": "SOW contemplates possible ground-mounted units.",
                    }
                ],
            },
        ],
    }


def mock_dossier_small():
    return {
        "bottom_line": (
            "Mechanical permit is typically required for furnace replacement. "
            "Electrical scope and any flue/combustion-air changes remain to be confirmed."
        ),
        "bottom_line_evidence": ["E1"],
        "research_completeness": {
            "status": "PARTIAL",
            "reason": "Core mechanical trigger visible; local permit form and electrical scope still open.",
            "critical_missing": ["Exact electrical scope", "Local mechanical permit application"],
        },
        "jurisdiction": {
            "status": "CONDITIONAL",
            "county": "Example County",
            "city": "Small Town",
            "ahj": "Local Building Department",
            "evidence": [],
        },
        "codes": [{"name": "Current state mechanical code", "status": "CONDITIONAL", "evidence": []}],
        "evidence": [
            {
                "id": "E1",
                "title": "Mechanical code applicability for furnace replacement",
                "url": "https://example.com",
                "authority": "state",
                "discipline": "Mechanical",
                "proposition_type": "APPLICABILITY",
                "source_type": "code",
                "retrieval_note": "General applicability.",
                "rule": "Replacement of a fuel-fired furnace is subject to the mechanical code.",
            }
        ],
        "disciplines": [
            {
                "type": "Mechanical",
                "applicability": {
                    "rule": "Furnace replacement is subject to the mechanical code.",
                    "fact": {"statement": "Replace existing gas furnace.", "source": "USER_PROVIDED"},
                    "determination": "applies",
                    "relationship": "direct",
                    "evidence": ["E1"],
                    "missing": "",
                },
                "permit": "CONDITIONAL",
                "permit_finding": "Mechanical permit expected; local permit-specific evidence not yet retrieved.",
                "permit_basis": "NOT_ESTABLISHED",
                "permit_evidence": [],
                "pathway": "CONDITIONAL",
                "pathway_finding": "Pathway not yet established.",
                "pathway_basis": "NOT_ESTABLISHED",
                "pathway_evidence": [],
                "missing": ["Whether flue or combustion air is altered"],
                "reopen": [],
                "actionable_questions": [
                    "Does the local AHJ allow over-the-counter permits for like-for-like furnace replacement?",
                ],
                "potential_issues": [],
            },
            {
                "type": "Electrical",
                "applicability": {
                    "rule": "Associated electrical work may require a permit.",
                    "fact": {"statement": "Electrical scope not stated.", "source": "USER_PROVIDED"},
                    "determination": "cannot_determine",
                    "relationship": "conditional",
                    "evidence": [],
                    "missing": "Whether disconnect, circuit, or wiring is modified.",
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
                    "Will the furnace replacement include a new disconnect or circuit changes?",
                ],
                "potential_issues": [],
            },
        ],
    }


# ============================================================
# LIVE GEMINI PATH (simplified but functional)
# ============================================================
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
            return json.loads(text[start : end + 1])
        raise ValueError("No valid JSON found")


def call_gemini(prompt: str) -> dict:
    """Minimal live call. Returns {data, error, msg, debug}."""
    debug = {"status": "started", "model": "gemini-2.0-flash"}
    if not GEMINI_KEY:
        return {"data": None, "error": True, "msg": "GEMINI_KEY missing.", "debug": debug}
    if not GENAI_AVAILABLE:
        return {"data": None, "error": True, "msg": "google-genai package not installed.", "debug": debug}

    try:
        client = genai.Client(api_key=GEMINI_KEY)
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                max_output_tokens=8192,
                response_mime_type="application/json",
            ),
        )
        text = getattr(response, "text", None) or ""
        data = extract_json(text)
        debug["status"] = "success"
        return {"data": data, "error": False, "msg": "", "debug": debug}
    except Exception as e:
        debug["status"] = "error"
        debug["exception"] = str(e)
        return {"data": None, "error": True, "msg": f"Gemini error: {str(e)[:300]}", "debug": debug}


def build_research_prompt(state, address, project_date, ptype, bclass, existing, sow):
    return f"""
You are an expert AHJ research analyst. Return ONLY valid JSON. No markdown or commentary.

PROJECT
State: {state}
Address: {address}
Date: {project_date}
Type: {ptype}
Building class: {bclass}
Existing entitlements: {existing or "None stated"}
Scope of Work:
{sow}

RESEARCH RULES (strict)
1. Jurisdiction must be based on parcel-level evidence. Postal city alone is NOT enough. If the parcel may be unincorporated, say so and set status CONDITIONAL.
2. Never invent permit requirements. A code applicability statement is NOT a permit requirement.
3. Use these exact proposition_type values for evidence: JURISDICTION, CODE_CURRENCY, APPLICABILITY, PERMIT_REQUIREMENT, PERMIT_EXEMPTION, REVIEW_REQUIREMENT, PATHWAY, THRESHOLD, ENTITLEMENT, OTHER.
4. permit = VERIFIED_REQUIRED only when you have real PERMIT_REQUIREMENT evidence whose rule text itself states that a permit is required.
5. Otherwise use CONDITIONAL or UNKNOWN and explain what is still missing.
6. Keep findings short. Prefer honesty over completeness.

Return this exact JSON shape:
{{
  "bottom_line": "2-3 short sentences",
  "bottom_line_evidence": ["E1"],
  "research_completeness": {{
    "status": "PARTIAL",
    "reason": "short reason",
    "critical_missing": ["item"]
  }},
  "jurisdiction": {{
    "status": "CONDITIONAL",
    "county": "",
    "city": "",
    "ahj": "",
    "evidence": ["E1"]
  }},
  "codes": [{{"name": "", "status": "CONDITIONAL", "evidence": []}}],
  "evidence": [{{
    "id": "E1",
    "title": "",
    "url": "",
    "authority": "state|county|city|other",
    "discipline": "",
    "proposition_type": "APPLICABILITY",
    "source_type": "code",
    "retrieval_note": "",
    "rule": "short proposition actually supported by the source"
  }}],
  "disciplines": [{{
    "type": "Mechanical",
    "applicability": {{
      "rule": "",
      "fact": {{"statement": "", "source": "USER_PROVIDED"}},
      "determination": "applies|cannot_determine|does_not_apply",
      "relationship": "direct|conditional",
      "evidence": [],
      "missing": ""
    }},
    "permit": "CONDITIONAL",
    "permit_finding": "",
    "permit_basis": "NOT_ESTABLISHED",
    "permit_evidence": [],
    "pathway": "CONDITIONAL",
    "pathway_finding": "",
    "pathway_basis": "NOT_ESTABLISHED",
    "pathway_evidence": [],
    "missing": [],
    "reopen": [],
    "actionable_questions": ["specific question"],
    "potential_issues": [{{"issue": "", "why": ""}}]
  }}]
}}
"""


# ============================================================
# UI
# ============================================================
if "report_data" not in st.session_state:
    st.session_state.report_data = None
if "debug_log" not in st.session_state:
    st.session_state.debug_log = {}
if "error_msg" not in st.session_state:
    st.session_state.error_msg = None
if "run_status" not in st.session_state:
    st.session_state.run_status = "Ready"

st.title("🏛️ AHJ Research Assistant")
st.caption("Evidence-contract architecture · Deterministic permit engine · Mock + live Gemini")

with st.sidebar:
    st.markdown("### Mode")
    mock_mode = st.toggle("🛡️ Mock Mode (recommended for testing)", value=True)
    st.markdown("---")
    st.markdown("**Quick presets**")
    if st.button("Large remodel (La Puente style)"):
        st.session_state["preset"] = "large"
    if st.button("Small furnace replacement"):
        st.session_state["preset"] = "small"

# Defaults / presets
preset = st.session_state.get("preset", "large")
if preset == "small":
    default_state = "Montana"
    default_address = "123 Main St, Small Town, MT"
    default_sow = "Replace existing gas furnace with like-for-like unit. No ductwork changes. Electrical scope unknown."
    default_class = "Residential (1-2 Family)"
else:
    default_state = "California"
    default_address = "13431 Temple Ave, La Puente, CA"
    default_sow = """EXTERIOR: structural repairs per discovery, window infill, stucco/siding, full exterior paint, roof cover replacement, possible ground-mounted HVAC.
INTERIOR: accessibility adjustments to lobby, bathrooms, platform; full paint; new ceilings.
ELECTRICAL: replace panel, branch circuits, fixtures.
PLUMBING: replace fixtures and water heaters; accessibility-related drainage work.
MECHANICAL: replace HVAC, prefer ground-mount if possible; replace ductwork.
SITE: parking re-stripe with ADA stalls, sidewalk accessibility work, CMU wall repairs.
Many items left as "repair per discovery" or "as required by code"."""
    default_class = "Assembly"

st.header("1. Project Metadata")
c1, c2 = st.columns(2)
with c1:
    state = st.selectbox("State", STATE_OPTIONS, index=STATE_OPTIONS.index(default_state))
    address = st.text_input("Project Address", default_address)
    project_date = st.date_input("Permit / Construction Date", date.today())
with c2:
    ptype = st.selectbox("Project Type", PROJECT_TYPES, index=1)
    bclass = st.selectbox("Building / Occupancy Class", BUILDING_CLASSES, index=BUILDING_CLASSES.index(default_class))
    existing = st.text_input("Existing Entitlements (optional)", placeholder="CUP, site plan, variance…")

st.header("2. Scope of Work")
sow = st.text_area("Paste complete Scope of Work", height=200, value=default_sow)

st.header("3. Run Research")
if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    st.session_state.error_msg = None
    st.session_state.report_data = None
    st.session_state.run_status = "Running…"

    if mock_mode:
        with st.spinner("Mock research + deterministic contract…"):
            raw = mock_dossier_large(address) if preset == "large" or "la puente" in _norm(address) else mock_dossier_small()
            data, errors = run_deterministic_contract(raw, address, state)
            st.session_state.report_data = data
            st.session_state.debug_log = {"mode": "mock", "validation_errors": errors}
            st.session_state.run_status = "Complete (mock)"
            if errors:
                st.session_state.error_msg = "Deterministic contract reported issues – see debug."
    else:
        with st.spinner("Calling Gemini (live)…"):
            prompt = build_research_prompt(state, address, project_date, ptype, bclass, existing, sow)
            result = call_gemini(prompt)
            st.session_state.debug_log = result.get("debug", {})
            if result["error"]:
                st.session_state.error_msg = result["msg"]
                st.session_state.run_status = "Failed"
            else:
                data, errors = run_deterministic_contract(result["data"], address, state)
                st.session_state.report_data = data
                st.session_state.debug_log["validation_errors"] = errors
                st.session_state.run_status = "Complete (live)"
                if errors:
                    st.session_state.error_msg = "Deterministic contract reported issues – see debug."

st.info(f"Status: **{st.session_state.run_status}**")

if st.session_state.error_msg:
    st.error(st.session_state.error_msg)

with st.expander("🐛 Debug", expanded=bool(st.session_state.error_msg)):
    st.json(st.session_state.debug_log)

# ============================================================
# RESULTS
# ============================================================
if st.session_state.report_data:
    data = st.session_state.report_data
    st.divider()
    st.header("4. Research Dossier")

    st.success(f"**Bottom Line**  \n{data.get('bottom_line', 'N/A')}")

    comp = data.get("research_completeness") or {}
    if comp:
        st.warning(f"Completeness: **{comp.get('status')}** — {comp.get('reason', '')}")
        missing = comp.get("critical_missing") or []
        if missing:
            st.markdown("**Still needed:**")
            for m in missing:
                st.markdown(f"- {m}")

    st.subheader("📍 Jurisdiction")
    jur = data.get("jurisdiction") or {}
    st.write(f"**Status:** {jur.get('status', 'N/A')}")
    st.write(f"**County:** {jur.get('county', '—')}")
    st.write(f"**City:** {jur.get('city', '—')}")
    st.write(f"**AHJ:** {jur.get('ahj', '—')}")
    if jur.get("validation_note"):
        st.caption(jur["validation_note"])

    st.subheader("📋 Disciplines")
    for item in data.get("disciplines") or []:
        permit = str(item.get("permit", "UNKNOWN")).upper()
        icon = {
            "VERIFIED_REQUIRED": "🟢",
            "CONDITIONAL": "🟠",
            "UNKNOWN": "🟡",
            "NOT_APPLICABLE": "⚪",
        }.get(permit, "🟡")
        with st.expander(f"{icon} {item.get('type', 'Unknown')} — {permit}", expanded=False):
            st.markdown(f"**Permit:** {item.get('permit_finding', '—')}")
            st.markdown(f"**Pathway:** {item.get('pathway_finding', '—')}")
            qs = item.get("actionable_questions") or []
            if qs:
                st.markdown("**Questions to resolve next**")
                for q in qs:
                    st.markdown(f"- {q}")
            leads = item.get("potential_issues") or []
            if leads:
                st.markdown("**Worth checking (research lead)**")
                for lead in leads:
                    if isinstance(lead, dict):
                        st.markdown(f"- {lead.get('issue', '')}")
                        if lead.get("why"):
                            st.caption(lead["why"])

    st.header("5. Export")
    json_str = json.dumps(
        {"project": {"state": state, "address": address, "date": str(project_date)}, "dossier": data},
        indent=2,
    )
    st.download_button(
        "💾 Download JSON",
        data=json_str,
        file_name="AHJ_Dossier.json",
        mime="application/json",
        use_container_width=True,
    )
