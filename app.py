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
from docx.shared import Pt

st.set_page_config(page_title="AHJ Research Assistant v25.1", page_icon="🏛️", layout="wide")

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

GEMINI_KEY = os.getenv("GEMINI_KEY") or st.secrets.get("GEMINI_KEY", "")
PROMPT_VERSION = "v25.1_regulatory_firewall"

# ============================================================
# HELPERS & VALIDATION
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
            return json.loads(text[start:end + 1])
        raise ValueError("No valid JSON found")

def validate_dossier(data):
    errors = []
    allowed_statuses = {"VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED", "UNKNOWN", "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED"}
    allowed_determinations = {"applies", "does_not_apply", "cannot_determine"}
    allowed_relationships = {"direct", "conditional", "not_established"}
    allowed_completeness = {"SUFFICIENT", "PARTIAL", "INSUFFICIENT"}
    
    # 2. Add stronger validation rules for evidence basis
    allowed_evidence_basis = {"DIRECT_EVIDENCE", "CONDITIONAL", "NOT_ESTABLISHED"}

    completeness = data.get("research_completeness", {})
    if completeness.get("status") not in allowed_completeness:
        errors.append("Research completeness must be SUFFICIENT, PARTIAL, or INSUFFICIENT.")
    if completeness.get("status") == "INSUFFICIENT" and not completeness.get("reason"):
        errors.append("INSUFFICIENT research completeness requires a reason.")

    evidence_items = data.get("evidence", [])
    evidence_by_id = {e.get("id"): e for e in evidence_items if e.get("id")}
    evidence_ids = set(evidence_by_id.keys())

    for ev in evidence_items:
        eid = ev.get("id", "Unknown")
        if not ev.get("title"): errors.append(f"{eid}: missing title.")
        if not ev.get("url"): errors.append(f"{eid}: missing URL.")
        if not ev.get("authority"): errors.append(f"{eid}: missing authority.")
        if not ev.get("discipline"): errors.append(f"{eid}: missing discipline.")
        if not ev.get("rule"): errors.append(f"{eid}: missing rule proposition.")

    jurisdiction = data.get("jurisdiction", {})
    for eid in jurisdiction.get("evidence", []):
        if eid not in evidence_ids: errors.append(f"Jurisdiction references nonexistent evidence: {eid}")
    if jurisdiction.get("status") == "VERIFIED" and not jurisdiction.get("evidence"):
        errors.append("Verified jurisdiction must have evidence.")

    for code in data.get("codes", []):
        for eid in code.get("evidence", []):
            if eid not in evidence_ids: errors.append(f"Code '{code.get('name')}' references nonexistent evidence: {eid}")
        if code.get("status") == "CURRENT" and not code.get("evidence"):
            errors.append(f"Current code '{code.get('name')}' must have evidence.")

    for eid in data.get("bottom_line_evidence", []):
        if eid not in evidence_ids: errors.append(f"Bottom Line references nonexistent evidence: {eid}")
    if not data.get("bottom_line_evidence"):
        errors.append("Bottom Line must reference supporting evidence IDs.")

    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        permit_basis = item.get("permit_basis")
        pathway_basis = item.get("pathway_basis")
        
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

        # 2. Validate permit and pathway basis
        if permit_basis not in allowed_evidence_basis:
            errors.append(f"{discipline}: invalid permit_basis '{permit_basis}'.")
        if pathway_basis not in allowed_evidence_basis:
            errors.append(f"{discipline}: invalid pathway_basis '{pathway_basis}'.")

        if fact_source not in FACT_SOURCES: errors.append(f"{discipline}: invalid fact source '{fact_source}'")
        if not fact_statement: errors.append(f"{discipline}: applicability fact is missing.")

        if fact_source in {"INFERRED", "UNKNOWN"} and determination == "does_not_apply":
            errors.append(f"{discipline}: cannot establish does_not_apply from inferred/unknown fact.")
        if fact_source == "INFERRED" and relationship == "direct":
            errors.append(f"{discipline}: inferred fact cannot support direct relationship.")

        for eid in app.get("evidence", []):
            if eid not in evidence_ids: errors.append(f"{discipline}: applicability references nonexistent evidence '{eid}'")
        for eid in item.get("permit_evidence", []):
            if eid not in evidence_ids: errors.append(f"{discipline}: permit references nonexistent evidence '{eid}'")
        for eid in item.get("pathway_evidence", []):
            if eid not in evidence_ids: errors.append(f"{discipline}: pathway references nonexistent evidence '{eid}'")

        for eid in app.get("evidence", []):
            evidence = evidence_by_id.get(eid)
            if evidence:
                ev_disc = (evidence.get("discipline") or "").lower()
                if ev_disc and discipline.lower() and ev_disc != discipline.lower() and relationship == "direct":
                    errors.append(f"{discipline}: direct applicability relies on {ev_disc} evidence {eid}. Explicit cross-discipline authority required.")

        # 2. Permit / pathway evidence-basis enforcement
        if permit == "VERIFIED_REQUIRED":
            if permit_basis != "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: VERIFIED_REQUIRED permit must have permit_basis=DIRECT_EVIDENCE.")
            if not item.get("permit_evidence"):
                errors.append(f"{discipline}: VERIFIED_REQUIRED permit requires permit-specific evidence.")
        if pathway == "VERIFIED_REQUIRED":
            if pathway_basis != "DIRECT_EVIDENCE":
                errors.append(f"{discipline}: VERIFIED_REQUIRED pathway must have pathway_basis=DIRECT_EVIDENCE.")
            if not item.get("pathway_evidence"):
                errors.append(f"{discipline}: VERIFIED_REQUIRED pathway requires pathway-specific evidence.")
        if permit_basis == "DIRECT_EVIDENCE" and not item.get("permit_evidence"):
            errors.append(f"{discipline}: DIRECT_EVIDENCE permit basis requires permit_evidence.")
        if pathway_basis == "DIRECT_EVIDENCE" and not item.get("pathway_evidence"):
            errors.append(f"{discipline}: DIRECT_EVIDENCE pathway basis requires pathway_evidence.")

        if relationship == "direct" and not app.get("rule"):
            errors.append(f"{discipline}: direct relationship requires a source rule.")
        if determination == "cannot_determine" and not app.get("missing"):
            errors.append(f"{discipline}: cannot_determine requires missing facts.")
        if app.get("missing") and determination == "does_not_apply":
            errors.append(f"{discipline}: unresolved facts present, but determination is does_not_apply.")
        if permit == "NOT_CURRENTLY_TRIGGERED" and determination == "does_not_apply" and app.get("missing"):
            errors.append(f"{discipline}: NOT_CURRENTLY_TRIGGERED cannot be paired with does_not_apply when missing facts exist.")
        
        # 4. Fix Structural/NOT_APPLICABLE logic specifically
        if permit == "NOT_APPLICABLE":
            if determination != "does_not_apply":
                errors.append(f"{discipline}: NOT_APPLICABLE requires determination=does_not_apply.")
            if relationship != "direct":
                errors.append(f"{discipline}: NOT_APPLICABLE requires a direct applicability determination.")
            if not app.get("evidence"):
                errors.append(f"{discipline}: NOT_APPLICABLE requires authoritative applicability evidence.")
            if app.get("missing"):
                errors.append(f"{discipline}: NOT_APPLICABLE cannot have unresolved applicability facts.")

        if relationship == "not_established":
            finding = (item.get("permit_finding") or "").lower()
            dangerous_phrases = ["is required", "requires", "shall require", "automatically triggers", "therefore requires", "must obtain"]
            for phrase in dangerous_phrases:
                if phrase in finding:
                    errors.append(f"{discipline}: relationship is not_established but permit_finding claims downstream requirement.")
                    break

        # 4. Strengthen the negative-conclusion validator
        finding_text = " ".join([str(item.get("permit_finding") or ""), str(item.get("pathway_finding") or "")]).lower()
        unsupported_negative_claims = [
            "no permit required", "no permit is required", "permit is not required", "permit not required", "does not require a permit", "does not trigger a permit",
            "no structural permit", "no electrical permit", "no mechanical permit", "no planning permit", "no land use permit",
            "no plan review", "plan review is not required", "plan review not required", "does not trigger review",
            "no cup amendment", "cup amendment is not required", "cup modification is not required", "no land use approval", "land use approval is not required",
            "no separate energy permit", "separate energy permit is not required", "energy permit is not required",
        ]
        if any(phrase in finding_text for phrase in unsupported_negative_claims):
            if not item.get("permit_evidence") and not item.get("pathway_evidence"):
                errors.append(f"{discipline}: negative permit/pathway conclusion lacks permit-specific or pathway-specific evidence.")

        # 6. Validator for "full plan review" claims
        review_claims = [
            "full plan review", "plan review is not triggered", "plan review not triggered",
            "trade permit without plan review", "no plan review", "administrative review",
            "trade permit pathway", "over-the-counter",
        ]
        if any(phrase in finding_text for phrase in review_claims):
            if not item.get("pathway_evidence"):
                errors.append(f"{discipline}: specific review-pathway claim lacks pathway-specific evidence.")

        # 3. Add a hard threshold firewall
        finding_text_all = " ".join([
            str(item.get("permit_finding") or ""),
            str(item.get("pathway_finding") or ""),
            str(app.get("rule") or ""),
        ]).lower()

        threshold_markers = [
            "lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "feet", "foot", 
            "percent", "%", "section", "chapter", "threshold", "exemption", "exception", 
            "over-the-counter", "minor label",
        ]

        has_threshold_claim = any(marker in finding_text_all for marker in threshold_markers)
        if has_threshold_claim:
            if not item.get("permit_evidence") and not item.get("pathway_evidence"):
                errors.append(f"{discipline}: regulatory threshold, exemption, section, or pathway claim lacks specific supporting evidence.")

        if pathway == "VERIFIED_REQUIRED" and relationship == "not_established":
            errors.append(f"{discipline}: pathway cannot be verified when relationship is not_established.")

    return errors

def validate_bottom_line(data):
    errors = []
    bottom_line = (data.get("bottom_line") or "").lower()
    
    # 13. Catch Bottom Line overreach (negative phrases)
    unsupported_phrases = [
        "no permit required", "no permit is required", "permit is not required", "permit not required",
        "no plan review", "plan review is not required", "does not trigger a permit", 
        "does not require a permit", "does not trigger review",
    ]
    if any(phrase in bottom_line for phrase in unsupported_phrases):
        errors.append("Bottom Line contains a negative permit/review conclusion that requires explicit supporting evidence.")

    # 15. Bottom Line numeric / pathway claim protection
    bottom_line_regulatory_markers = [
        "lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "section", "chapter",
        "threshold", "over-the-counter", "minor label", "full plan review", "trade permit",
        "administrative review", "cup amendment", "cup modification", "energy permit",
    ]
    if any(marker in bottom_line for marker in bottom_line_regulatory_markers):
        if not data.get("bottom_line_evidence"):
            errors.append("Bottom Line contains a specific regulatory claim without supporting evidence IDs.")

    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        app = item.get("applicability", {})
        determination = app.get("determination")

        unresolved = (permit in {"CONDITIONAL", "UNKNOWN", "NOT_CURRENTLY_TRIGGERED"} or 
                      pathway in {"CONDITIONAL", "UNKNOWN", "NOT_CURRENTLY_TRIGGERED"} or 
                      determination == "cannot_determine")
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
# CACHING & API CALL (SEPARATED RETRY LOGIC)
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": 1}
    
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())],
            max_output_tokens=8192,
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

        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "error": True, "retry": False, "msg": "Failed to parse JSON.", "debug": debug_info}

        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            
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
            tools=[types.Tool(google_search=types.GoogleSearch())],
            max_output_tokens=8192,
        )

        response = client.models.generate_content(
            model="gemini-3.6-flash", 
            contents=retry_prompt,
            config=config,
        )
        
        if not response.candidates:
            return {"data": None, "error": True, "retry": False, "msg": "Retry returned no candidates.", "debug": debug_info}

        candidate = response.candidates[0]
        finish_reason = str(candidate.finish_reason)
        debug_info["finish_reason"] = finish_reason

        if finish_reason == "FinishReason.MAX_TOKENS":
            debug_info["error_type"] = "MAX_TOKENS_RETRY"
            return {"data": None, "error": True, "retry": False, "msg": "Model reached the output limit on the compact retry.", "debug": debug_info}

        if finish_reason and finish_reason != "FinishReason.STOP":
            return {"data": None, "error": True, "retry": False, "msg": f"Retry stopped early: {finish_reason}", "debug": debug_info}

        text = getattr(response, "text", None)
        if not text:
            try: text = response.candidates[0].content.parts[0].text
            except Exception: pass
                
        if not text:
            return {"data": None, "error": True, "retry": False, "msg": "Retry returned empty text.", "debug": debug_info}

        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["error_type"] = "JSON Parse Failed"
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:1000]
            return {"data": None, "error": True, "retry": False, "msg": "Retry produced invalid JSON.", "debug": debug_info}

        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            
        debug_info["status"] = "success"
        return {"data": data, "error": False, "retry": False, "debug": debug_info}
        
    except Exception as e:
        debug_info["exception"] = str(e)
        debug_info["error_type"] = "Python Exception"
        return {"data": None, "error": True, "retry": False, "msg": f"Retry error: {str(e)[:200]}", "debug": debug_info}

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state: st.session_state.error_msg = None

st.title("🏛️ AHJ Research Assistant v25.1")
st.caption("Regulatory inference firewall. Evidence-basis enforcement. Code currency tracking.")

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
    # 1. Fix the entitlement input
    existing_permit = st.text_input(
        "Existing Entitlements (Optional)",
        "",
        placeholder="e.g., CUP, variance, site plan, development agreement"
    )

# 15. Cleaner default SOW that explicitly states what is unknown
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
        # 11. Fix mock data to test new rules conservatively
        st.session_state.report_data = {
            "bottom_line": "Mechanical permit requirement is established. Specific review pathway remains conditional. Electrical scope is unknown. Structural and Planning determinations require retrieval of governing conditions and equipment specifications.",
            "bottom_line_evidence": ["E2", "E3", "E4"],
            "research_completeness": {"status": "PARTIAL", "reason": "Main regulatory framework established, but equipment specs and CUP conditions remain unresolved.", "critical_missing": ["Equipment cut sheets", "Existing CUP document"]},
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Jurisdiction", "source_type": "permit_page", "retrieval_note": "Confirms jurisdiction for unincorporated areas.", "rule": "Jurisdiction for unincorporated areas requires county confirmation."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Mechanical", "source_type": "checklist", "retrieval_note": "Explicitly lists commercial HVAC replacement requirements.", "rule": "Mechanical replacement is regulated by the adopted mechanical code and requires a permit."},
                {"id": "E3", "title": "Oregon Energy Efficiency Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Energy", "source_type": "code", "retrieval_note": "Governs replacement equipment efficiency.", "rule": "Replacement mechanical equipment must comply with current energy efficiency standards."},
                {"id": "E4", "title": "Oregon Electrical Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Electrical", "source_type": "code", "retrieval_note": "Defines when electrical modifications trigger permits.", "rule": "Electrical permit required for modification of branch circuits or disconnects."}
            ],
            "disciplines": [
                {
                    "type": "Mechanical",
                    "applicability": {"rule": "Mechanical replacement is regulated by the adopted mechanical code.", "fact": {"statement": "Replacing ground-level exterior commercial HVAC equipment.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E2"]},
                    "permit": "VERIFIED_REQUIRED", "permit_finding": "Mechanical permit requirement is established by the cited permit/code evidence.", "permit_basis": "DIRECT_EVIDENCE", "permit_evidence": ["E2"],
                    "pathway": "CONDITIONAL", "pathway_finding": "Specific review pathway remains unresolved because project-specific pathway criteria have not been verified.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [],
                    "missing": ["Project-specific review pathway criteria (e.g., unit weight, CFM)."], "reopen": ["Authoritative pathway criteria retrieved."]
                },
                {
                    "type": "Electrical",
                    "applicability": {"rule": "Electrical permit required for branch circuit/disconnect modification.", "fact": {"statement": "Electrical modifications to replacement unit are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Whether wiring/disconnect/breaker will be modified.", "relationship": "conditional", "evidence": ["E4"]},
                    "permit": "CONDITIONAL", "permit_finding": "Electrical permit consequence depends on whether wiring/disconnect/circuit work occurs.", "permit_basis": "CONDITIONAL", "permit_evidence": ["E4"],
                    "pathway": "CONDITIONAL", "pathway_finding": "Pathway cannot be determined until electrical scope is defined.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [],
                    "missing": ["Unit electrical specs (MCA, MOP, voltage).", "Scope of electrical changes."], "reopen": ["Modifying electrical disconnect, wiring, or breaker."]
                },
                {
                    "type": "Energy",
                    "applicability": {"rule": "Current energy code applies to replacement equipment.", "fact": {"statement": "Replacing HVAC unit; efficiency ratings unknown.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E3"]},
                    "permit": "CONDITIONAL", "permit_finding": "Energy code compliance applies, separate permit unconfirmed.", "permit_basis": "CONDITIONAL", "permit_evidence": [],
                    "pathway": "CONDITIONAL", "pathway_finding": "Compliance pathway depends on equipment specifications.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [],
                    "missing": ["Replacement equipment efficiency/specifications."], "reopen": []
                },
                {
                    "type": "Structural",
                    "applicability": {"rule": "Structural permits required for alterations or added loads.", "fact": {"statement": "Replacement unit weight and anchorage configuration are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Replacement unit operating weight and anchorage configuration.", "relationship": "not_established", "evidence": []},
                    "permit": "UNKNOWN", "permit_finding": "Permit consequence cannot be determined without equipment specifications.", "permit_basis": "NOT_ESTABLISHED", "permit_evidence": [],
                    "pathway": "UNKNOWN", "pathway_finding": "Review pathway cannot be established.", "pathway_basis": "NOT_ESTABLISHED", "pathway_evidence": [],
                    "missing": ["Replacement unit operating weight.", "Anchorage configuration."], "reopen": ["Rooftop mounting, suspended installation, or structural framing modifications occur."]
                },
                {
                    "type": "Planning / CUP",
                    "applicability": {"rule": "Work must comply with existing CUP conditions.", "fact": {"statement": "Parcel operates under existing CUP; actual conditions not retrieved.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Actual CUP conditions governing exterior equipment.", "relationship": "conditional", "evidence": ["E1"]},
                    "permit": "CONDITIONAL", "permit_finding": "Existing CUP identified, but governing conditions have not been reviewed.", "permit_basis": "CONDITIONAL", "permit_evidence": [],
                    "pathway": "CONDITIONAL", "pathway_finding": "Final land-use determination conditional on review of existing CUP conditions.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [],
                    "missing": ["Actual CUP conditions governing exterior equipment."], "reopen": ["Relocation, footprint expansion, screening changes, noise increases, or site work occur."]
                }
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
You are an expert AHJ research analyst. Return ONLY a valid JSON object. No conversational text, markdown, or explanations outside the JSON.

PROJECT:
State: {state} | Address: {address} | Date: {project_date}
Type: {ptype} | Class: {bclass} (USER-PROVIDED) | Entitlements: {existing_permit}
SCOPE: {sow_text}

OUTPUT COMPRESSION RULE: Keep the final JSON concise and information-dense. Avoid explanatory prose, repetition, or conclusions that duplicate the matrix. Target lengths (flexible, prioritize correctness): rule: 10-30 words, project fact: 5-20 words, permit/pathway finding: 10-30 words, missing/reopen items: 3-15 words.

DISCIPLINE DISCOVERY:
Determine disciplines dynamically from: 1) project type; 2) explicit SOW facts; 3) jurisdiction; 4) adopted code scope; 5) permit authority requirements; 6) land-use/entitlement information; 7) other authoritative applicability rules. Do not use a fixed discipline list. Do not impose a maximum number of disciplines. Do not omit a discipline merely because the SOW does not mention it. However, do not manufacture a requirement merely because a discipline exists. Only conclude NOT_APPLICABLE when authoritative evidence supports that determination for the known project facts.

REGULATORY CONCLUSION LEVELS:
Treat each regulatory conclusion as five separate levels:
LEVEL 1 — SOURCE FACT: What exactly does the retrieved authoritative source state?
LEVEL 2 — PROJECT FACT: What exactly is known about this project from the user or retrieved project records?
LEVEL 3 — APPLICABILITY: Does the source rule apply to the known project facts?
LEVEL 4 — PERMIT CONSEQUENCE: Does authoritative evidence explicitly establish a permit requirement or exemption?
LEVEL 5 — REVIEW PATHWAY: Does authoritative evidence explicitly establish the review pathway, such as trade permit, plan review, administrative review, over-the-counter review, minor label, or another specific process?

CRITICAL FIREWALL:
Level 1 + Level 2 does NOT automatically establish Level 4 or Level 5. A code threshold does not automatically establish a permit pathway. A code applicability rule does not automatically establish a permit requirement. A permit requirement does not automatically establish a particular review pathway. An exemption does not automatically establish that the project qualifies for the exemption. A definition does not automatically establish a permit consequence.
For every permit conclusion, ask: "What retrieved source explicitly establishes this permit consequence?"
For every pathway conclusion, ask: "What retrieved source explicitly establishes this pathway?"
If that source cannot be identified: do not mark the conclusion VERIFIED_REQUIRED; do not invent the pathway from general code knowledge; use CONDITIONAL or UNKNOWN; explain what evidence is missing.

PERMIT BASIS: Set permit_basis="DIRECT_EVIDENCE" only when permit_evidence directly supports the stated permit conclusion. Otherwise use CONDITIONAL (when consequence depends on unresolved project facts) or NOT_ESTABLISHED (when available evidence does not establish the consequence).
PATHWAY BASIS: Set pathway_basis="DIRECT_EVIDENCE" only when pathway_evidence directly supports the stated pathway. Otherwise use CONDITIONAL or NOT_ESTABLISHED.
Never use the words "over-the-counter", "minor label", "full plan review", "trade permit", "administrative review", or similar pathway terms unless pathway-specific authoritative evidence was retrieved.

EXISTING ENTITLEMENTS / CUP RULE:
If the project identifies an existing CUP, variance, site plan, development agreement, land-use approval, or other governing entitlement: 1) Do not assume its conditions. 2) Do not infer that like-for-like work is exempt. 3) Do not conclude that an amendment is unnecessary unless the actual governing conditions or authoritative local rule supports that conclusion. 4) If the governing document has not been retrieved and could affect the determination, mark Planning / CUP applicability as cannot_determine or the permit consequence as CONDITIONAL. 5) Identify the exact governing document that should be retrieved. 6) "Same location" is a project fact, not proof of land-use exemption.

ELECTRICAL SCOPE:
Do not infer electrical work from HVAC replacement. Do not infer that an existing disconnect will be replaced, the circuit is adequate, or that MCA/MOP, voltage, phase, breaker size, conductor size, disconnect type, or wiring will remain unchanged. If electrical scope is not stated or documented: identify it as UNKNOWN; do not convert the uncertainty into "no electrical permit"; use CONDITIONAL where the permit consequence depends on whether electrical work occurs.

ENERGY RULE:
Do not infer the permit or review pathway from energy-code applicability. If the energy code applies, state that separately. Only state that energy compliance is reviewed through a particular permit or application process when authoritative evidence establishes that pathway. Do not invent a "separate energy permit" or conclude that no separate energy permit exists without authoritative evidence. Missing efficiency data affects compliance determination, not necessarily code applicability.

SCOPE DETAIL VARIABILITY:
The SOW may be extremely brief or highly detailed. Do NOT require a minimum SOW length. Do NOT assume that a short SOW means the project is simple. Do NOT assume facts that are normally included in a detailed construction scope. Research all disciplines that are reasonably implicated. When project facts are missing: continue researching the applicable regulatory framework; identify the missing project facts; distinguish applicability from permit consequence; distinguish permit consequence from review pathway; use CONDITIONAL, UNKNOWN, or NOT_CURRENTLY_TRIGGERED as appropriate. A sparse SOW should produce a useful screening dossier with clearly identified unknowns, not a weaker research effort. Never treat missing information as evidence that a discipline does not apply.

PROJECT FACT PROVENANCE:
The SOW and explicit project metadata supplied by the user are sources of USER_PROVIDED project facts. A fact is USER_PROVIDED only if explicitly stated. Absence of a fact is UNKNOWN, not "does not exist." Do not transform typical construction practice, project type, equipment type, or the analyst's expectations into project facts. For example, these remain UNKNOWN unless explicitly stated or retrieved: electrical reconnection, disconnect replacement, breaker replacement, circuit adequacy, voltage, phase, MCA/MOP, equipment weight, anchorage, structural capacity, screening, sound level, setback compliance, CUP conditions, zoning compliance, permit pathway, review pathway.

EVIDENCE QUALITY & CODE CURRENCY:
Every evidence item must support a specific proposition. Prefer: actual code sections, official permit requirements, official application/checklist pages, official ordinances, official land-use decisions, official CUP/entitlement documents, official code interpretations. Do not use a generic agency homepage as evidence for a specific permit requirement unless that page actually states the requirement.
For every code marked CURRENT: evidence must come from the authoritative adopting jurisdiction; the evidence must establish the code edition and effective/current status; do not use a generic code summary page when the actual adoption/currentness page is available. If the project date falls within a code transition or phase-in period, describe that explicitly rather than selecting an edition silently.

DISCIPLINE INDEPENDENCE: Evidence for Discipline A cannot establish requirements for Discipline B without explicit cross-discipline authority.

BOTTOM LINE EVIDENCE REQUIREMENT:
Every material statement in Bottom Line must be traceable to the matrix. If a sentence contains: a permit requirement, a permit exemption, a numeric threshold, a code edition, a review pathway, a land-use conclusion, a CUP conclusion, or an energy compliance pathway, then the underlying discipline must contain evidence supporting that exact proposition. Do not use Bottom Line as a place to resolve uncertainty that remains unresolved in the matrix. When evidence supports only applicability, say applicability is established and identify permit/pathway status separately. When a permit is established but pathway is unresolved, say the permit is established and the pathway remains conditional. When important facts are missing, preserve that uncertainty.

RESEARCH COMPLETENESS:
Set SUFFICIENT only when the retrieved authoritative evidence is adequate to support the material conclusions. Set PARTIAL when the main regulatory framework is established but one or more material facts or governing documents remain unresolved. Set INSUFFICIENT when jurisdiction, governing code, permit authority, or material regulatory requirements cannot be established. Do not use "SUFFICIENT" merely because the model is confident.

JSON SCHEMA:
{{
  "bottom_line": "3-5 sentences maximum.",
  "bottom_line_evidence": ["E1"],
  "research_completeness": {{
    "status": "SUFFICIENT|PARTIAL|INSUFFICIENT",
    "reason": "string",
    "critical_missing": ["string"]
  }},
  "jurisdiction": {{
    "status": "VERIFIED or CONDITIONAL",
    "county": "string",
    "city": "string",
    "ahj": "string",
    "evidence": ["E1"]
  }},
  "codes": [{{"name": "string", "status": "CURRENT or CONDITIONAL", "evidence": ["E2"]}}],
  "evidence": [{{
    "id": "E1", "title": "string", "url": "string", 
    "authority": "state|county|city|federal|tribal|other",
    "discipline": "Mechanical|Electrical|Structural|Planning|Energy|Jurisdiction|etc",
    "source_type": "code|ordinance|permit_page|checklist|application|interpretation|entitlement|other",
    "retrieval_note": "Why this source is relevant to the proposition.",
    "rule": "Specific proposition established by this source."
  }}],
  "disciplines": [{{
    "type": "string",
    "applicability": {{
      "rule": "string",
      "fact": {{"statement": "string", "source": "USER_PROVIDED|RETRIEVED_RECORD|AUTHORITATIVE_SOURCE|INFERRED|UNKNOWN"}},
      "determination": "applies|does_not_apply|cannot_determine",
      "missing": "string",
      "relationship": "direct|conditional|not_established",
      "evidence": ["E1"]
    }},
    "permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "permit_finding": "string",
    "permit_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
    "permit_evidence": ["E2"],
    "pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "pathway_finding": "string",
    "pathway_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
    "pathway_evidence": ["E3"],
    "missing": ["string"],
    "reopen": ["string"]
  }}]
}}
"""
                result = cached_gemini_call(prompt_hash, prompt)

                if result.get("retry"):
                    retry_prompt = prompt + """

COMPACT RETRY — PRESERVE RESEARCH QUALITY
The previous response exceeded the output budget.
Do NOT perform less research. Do NOT remove disciplines. Do NOT remove authoritative evidence. Do NOT replace evidence with model knowledge.
Instead: Keep each material discipline. Keep applicability, permit, and pathway separate. Keep only the strongest proposition-specific evidence. Remove repetition. Keep each finding concise. Keep missing facts concise. Keep reopen conditions concise. Keep Bottom Line to 3-5 sentences.
CRITICAL: A compact response must be less verbose, NOT less rigorous. Do not convert uncertain conclusions into definitive ones merely to save tokens.
"""
                    retry_result = cached_gemini_retry(prompt_hash + "_retry", retry_prompt)
                    debug_combined = {
                        "first_attempt": result.get("debug", {}),
                        "retry_attempt": retry_result.get("debug", {}),
                    }
                    st.session_state.debug_log = debug_combined
                    result = retry_result
                else:
                    st.session_state.debug_log = result.get("debug", {})

                if result["error"]:
                    st.session_state.error_msg = result["msg"]
                else:
                    st.session_state.report_data = result["data"]

if st.session_state.error_msg:
    st.error(f"❌ {st.session_state.error_msg}")

with st.expander("🐛 API Debug Log", expanded=False):
    st.json(st.session_state.debug_log)

# ============================================================
# RESULTS DISPLAY
# ============================================================
if st.session_state.report_data:
    data = st.session_state.report_data
    st.divider()
    st.header("4. Research Dossier")
    
    if "validation_errors" in st.session_state.debug_log:
        st.warning("⚠️ Schema Validation Warnings: " + " | ".join(st.session_state.debug_log["validation_errors"]))

    st.info(f"**Bottom Line:** {data.get('bottom_line', 'N/A')}")
    bl_ev = data.get("bottom_line_evidence", [])
    if bl_ev: st.caption(f"Bottom Line Evidence IDs: {', '.join(bl_ev)}")

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
            st.markdown("**Critical Missing Information:**")
            for item in critical_missing:
                st.markdown(f"- {item}")

    st.subheader("📍 Jurisdiction Determination")
    jur = data.get("jurisdiction") or {}
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**County:** {jur.get('county', 'Unknown')}")
        st.write(f"**City:** {jur.get('city', 'Unknown')}")
        st.write(f"**Building AHJ:** {jur.get('ahj', 'Unknown')}")
    with col2:
        ev_dict = {ev["id"]: ev for ev in data.get("evidence", [])}
        for eid in jur.get("evidence", []):
            if eid in ev_dict: st.write(f"**Source:** [{ev_dict[eid]['title']}]({ev_dict[eid]['url']})")

    # 13. Display code evidence, too
    st.subheader("📚 Applicable Codes")
    for code in (data.get("codes") or []):
        code_name = code.get("name", "Unknown")
        code_status = code.get("status", "N/A")
        st.markdown(f"- **{code_name}** ({code_status})")
        code_evidence = code.get("evidence", [])
        for eid in code_evidence:
            if eid in ev_dict:
                ev = ev_dict[eid]
                st.caption(f"  Evidence: {ev.get('title', eid)} — {ev.get('rule', 'N/A')}")

    st.subheader("📋 Permit & Review Matrix")
    status_map = {"VERIFIED_REQUIRED": "", "INFERRED": "🟡", "CONDITIONAL": "🟠", "UNKNOWN": "🔴", "NOT_APPLICABLE": "⚪", "NOT_CURRENTLY_TRIGGERED": "⚪", "USER_PROVIDED": ""}

    for item in data.get("disciplines", []):
        permit_emoji = status_map.get(item.get("permit", "UNKNOWN"), "")
        pathway_emoji = status_map.get(item.get("pathway", "UNKNOWN"), "")
        expander_title = f"{permit_emoji} {item.get('type', 'Unknown')} — Permit: {item.get('permit', 'UNKNOWN')} | Pathway: {pathway_emoji} {item.get('pathway', 'UNKNOWN')}"
        
        with st.expander(expander_title, expanded=False):
            app = item.get("applicability", {})
            fact = app.get("fact", {})
            fact_statement = fact.get("statement", "N/A") if isinstance(fact, dict) else str(fact)
            fact_source = fact.get("source", "UNKNOWN") if isinstance(fact, dict) else "UNKNOWN"
            
            st.markdown(f"**Applicability:** {app.get('determination', 'N/A').replace('_', ' ').title()}")
            st.markdown(f"**Source Rule:** {app.get('rule', 'N/A')}")
            st.markdown(f"**Project Fact:** {fact_statement} `[{fact_source}]`")
            st.markdown(f"**Relationship:** {app.get('relationship', 'N/A').replace('_', ' ').title()}")
            
            st.divider()
            st.markdown(f"**Permit:** {item.get('permit', 'N/A')}")
            st.markdown(f"**Permit Finding:** {item.get('permit_finding', 'N/A')}")
            # 12. Display the evidence basis in the UI
            permit_basis = item.get("permit_basis", "NOT_ESTABLISHED")
            st.caption(f"Permit evidence basis: {permit_basis}")
            
            st.markdown(f"**Pathway:** {item.get('pathway', 'N/A')}")
            st.markdown(f"**Pathway Finding:** {item.get('pathway_finding', 'N/A')}")
            pathway_basis = item.get("pathway_basis", "NOT_ESTABLISHED")
            st.caption(f"Pathway evidence basis: {pathway_basis}")
            
            st.divider()
            app_ev, permit_ev, pathway_ev = app.get("evidence", []), item.get("permit_evidence", []), item.get("pathway_evidence", [])
            
            if app_ev:
                st.write("**Applicability Evidence:**")
                for eid in app_ev:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})** `[{ev.get('authority', 'other').upper()}]` `[{ev.get('discipline', 'general').upper()}]`")
                        st.caption(f"  *Type:* {ev.get('source_type', 'N/A')} | *Note:* {ev.get('retrieval_note', 'N/A')}")
                        st.caption(f"  *Rule:* {ev.get('rule', 'N/A')}")
            if permit_ev:
                st.write("**Permit Evidence:**")
                for eid in permit_ev:
                    if eid in ev_dict: st.markdown(f"- **[{ev_dict[eid]['title']}]({ev_dict[eid]['url']})**")
            else: st.caption("*No permit-specific evidence retrieved.*")
            if pathway_ev:
                st.write("**Pathway Evidence:**")
                for eid in pathway_ev:
                    if eid in ev_dict: st.markdown(f"- **[{ev_dict[eid]['title']}]({ev_dict[eid]['url']})**")
            else: st.caption("*No pathway-specific evidence retrieved.*")
            
            missing = item.get("missing", [])
            if isinstance(missing, str): missing = [missing]
            if missing:
                st.markdown("**Missing Information:**")
                for fact in missing: st.markdown(f"- {fact}")
            reopen = item.get("reopen", [])
            if isinstance(reopen, str): reopen = [reopen]
            if reopen:
                st.markdown("**Reopen If:**")
                for condition in reopen: st.markdown(f"- {condition}")

    st.header("5. Export")
    col1, col2 = st.columns(2)
    with col1:
        doc = Document()
        doc.styles["Normal"].font.name = "Aptos"
        doc.add_heading("AHJ Research Dossier", 0)
        doc.add_paragraph(f"Project: {address} ({state})\nDate: {project_date}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
        doc.add_heading("Bottom Line", level=1)
        doc.add_paragraph(data.get("bottom_line", ""))
        doc.add_heading("Jurisdiction", level=1)
        doc.add_paragraph(f"Status: {jur.get('status', 'N/A')}\nCounty: {jur.get('county', 'N/A')}\nCity: {jur.get('city', 'N/A')}\nAHJ: {jur.get('ahj', 'N/A')}")
        doc.add_heading("Applicable Codes", level=1)
        for code in (data.get("codes") or []): doc.add_paragraph(f"{code.get('name')} ({code.get('status')})", style='List Bullet')
        doc.add_heading("Permit Matrix", level=1)
        for item in data.get("disciplines", []):
            doc.add_heading(f"{item.get('type')} - Permit: {item.get('permit')} | Pathway: {item.get('pathway')}", level=2)
            app = item.get("applicability", {})
            fact = app.get("fact", {})
            fact_statement = fact.get("statement", "") if isinstance(fact, dict) else str(fact)
            fact_source = fact.get("source", "UNKNOWN") if isinstance(fact, dict) else "UNKNOWN"
            doc.add_paragraph(f"Applicability: {app.get('determination', '').replace('_', ' ').title()}")
            doc.add_paragraph(f"Source Rule: {app.get('rule')}")
            doc.add_paragraph(f"Project Fact: {fact_statement} [{fact_source}]")
            doc.add_paragraph(f"Relationship: {app.get('relationship', '').replace('_', ' ').title()}")
            doc.add_paragraph(f"Permit Finding: {item.get('permit_finding')} (Basis: {item.get('permit_basis', 'NOT_ESTABLISHED')})")
            doc.add_paragraph(f"Pathway Finding: {item.get('pathway_finding')} (Basis: {item.get('pathway_basis', 'NOT_ESTABLISHED')})")
            missing = item.get("missing", [])
            if isinstance(missing, str): missing = [missing]
            if missing:
                doc.add_paragraph("Missing Information:", style='List Bullet')
                for fact in missing: doc.add_paragraph(f"  - {fact}")
            reopen = item.get("reopen", [])
            if isinstance(reopen, str): reopen = [reopen]
            if reopen:
                doc.add_paragraph("Reopen If:", style='List Bullet')
                for condition in reopen: doc.add_paragraph(f"  - {condition}")
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📄 Download Word Report", data=buf.getvalue(), file_name="AHJ_Dossier.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with col2:
        json_data = json.dumps({"project": {"state": state, "address": address, "date": str(project_date)}, "dossier": data}, indent=2)
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Dossier.json", mime="application/json", use_container_width=True)
