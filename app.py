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

st.set_page_config(page_title="AHJ Research Assistant v25.4", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v25.4_strict_fact_integrity"

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

    def discipline_family(value):
        value = value.lower()
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

    compatible_families = {
        frozenset(["building", "mechanical"]),
        frozenset(["building", "electrical"]),
        frozenset(["building", "structural"]),
        frozenset(["building", "energy"]),
        frozenset(["mechanical", "energy"]),
        frozenset(["planning", "building"]),
    }

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
        if permit_basis not in allowed_evidence_basis: errors.append(f"{discipline}: invalid permit_basis '{permit_basis}'.")
        if pathway_basis not in allowed_evidence_basis: errors.append(f"{discipline}: invalid pathway_basis '{pathway_basis}'.")
        if fact_source not in FACT_SOURCES: errors.append(f"{discipline}: invalid fact source '{fact_source}'")
        if not fact_statement: errors.append(f"{discipline}: applicability fact is missing.")

        # STRICT FACT INTEGRITY VALIDATION
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
                same_family = (ev_family == current_family)
                compatible = (frozenset([ev_family, current_family]) in compatible_families)
                if ev_family and current_family and not same_family and not compatible:
                    errors.append(f"{discipline}: direct applicability relies on {evidence.get('discipline', 'other')} evidence {eid}. Explicit cross-discipline authority required.")

        for eid in item.get("permit_evidence", []):
            if eid not in evidence_ids: errors.append(f"{discipline}: permit references nonexistent evidence '{eid}'")
        for eid in item.get("pathway_evidence", []):
            if eid not in evidence_ids: errors.append(f"{discipline}: pathway references nonexistent evidence '{eid}'")

        if permit == "VERIFIED_REQUIRED":
            if permit_basis != "DIRECT_EVIDENCE": errors.append(f"{discipline}: VERIFIED_REQUIRED permit must have permit_basis=DIRECT_EVIDENCE.")
            if not item.get("permit_evidence"): errors.append(f"{discipline}: VERIFIED_REQUIRED permit requires permit-specific evidence.")
        if pathway == "VERIFIED_REQUIRED":
            if pathway_basis != "DIRECT_EVIDENCE": errors.append(f"{discipline}: VERIFIED_REQUIRED pathway must have pathway_basis=DIRECT_EVIDENCE.")
            if not item.get("pathway_evidence"): errors.append(f"{discipline}: VERIFIED_REQUIRED pathway requires pathway-specific evidence.")
        if permit_basis == "DIRECT_EVIDENCE" and not item.get("permit_evidence"):
            errors.append(f"{discipline}: DIRECT_EVIDENCE permit basis requires permit_evidence.")
        if pathway_basis == "DIRECT_EVIDENCE" and not item.get("pathway_evidence"):
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
            finding = (item.get("permit_finding") or "").lower()
            for phrase in ["is required", "requires", "shall require", "automatically triggers", "therefore requires", "must obtain"]:
                if phrase in finding:
                    errors.append(f"{discipline}: relationship is not_established but permit_finding claims downstream requirement.")
                    break

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
                errors.append(f"{discipline}: negative permit/pathway conclusion lacks specific evidence.")

        review_claims = ["full plan review", "plan review is not triggered", "plan review not triggered", "trade permit without plan review", "no plan review", "administrative review", "trade permit pathway", "over-the-counter"]
        if any(phrase in finding_text for phrase in review_claims):
            if not item.get("pathway_evidence"):
                errors.append(f"{discipline}: specific review-pathway claim lacks pathway-specific evidence.")

        finding_text_all = " ".join([str(item.get("permit_finding") or ""), str(item.get("pathway_finding") or ""), str(app.get("rule") or "")]).lower()
        threshold_markers = ["lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "feet", "foot", "percent", "%", "section", "chapter", "threshold", "exemption", "exception", "over-the-counter", "minor label"]
        if any(marker in finding_text_all for marker in threshold_markers):
            if not item.get("permit_evidence") and not item.get("pathway_evidence"):
                errors.append(f"{discipline}: regulatory threshold/exemption/pathway claim lacks specific supporting evidence.")

        if pathway == "VERIFIED_REQUIRED" and relationship == "not_established":
            errors.append(f"{discipline}: pathway cannot be verified when relationship is not_established.")

    return errors

def validate_bottom_line(data):
    errors = []
    bottom_line = (data.get("bottom_line") or "").lower()
    
    unsupported_phrases = ["no permit required", "no permit is required", "permit is not required", "permit not required", "no plan review", "plan review is not required", "does not trigger a permit", "does not require a permit", "does not trigger review"]
    if any(phrase in bottom_line for phrase in unsupported_phrases):
        errors.append("Bottom Line contains a negative permit/review conclusion that requires explicit supporting evidence.")

    bottom_line_regulatory_markers = ["lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "section", "chapter", "threshold", "over-the-counter", "minor label", "full plan review", "trade permit", "administrative review", "cup amendment", "cup modification", "energy permit"]
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
# CACHING & API CALL
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": 1}
    
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            max_output_tokens=16384,
            thinking_config=types.ThinkingConfig(thinking_level="medium"),
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
            max_output_tokens=16384,
            thinking_config=types.ThinkingConfig(thinking_level="medium"),
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
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state: st.session_state.error_msg = None

st.title("🏛️ AHJ Research Assistant v25.4")
st.caption("16K generation ceiling. Medium reasoning. Evidence and regulatory consistency validation enabled.")

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
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Jurisdiction", "source_type": "permit_page", "retrieval_note": "Confirms jurisdiction for unincorporated areas.", "rule": "Jurisdiction for unincorporated areas requires county confirmation."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Mechanical", "source_type": "checklist", "retrieval_note": "Explicitly lists commercial HVAC replacement requirements.", "rule": "Mechanical replacement is regulated by the adopted mechanical code and requires a permit."},
                {"id": "E3", "title": "Oregon Energy Efficiency Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Energy", "source_type": "code", "retrieval_note": "Governs replacement equipment efficiency.", "rule": "Replacement mechanical equipment must comply with current energy efficiency standards."},
                {"id": "E4", "title": "Oregon Electrical Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Electrical", "source_type": "code", "retrieval_note": "Defines when electrical modifications trigger permits.", "rule": "Electrical permit required for modification of branch circuits or disconnects."}
            ],
            "disciplines": [
                {"type": "Mechanical", "applicability": {"rule": "Mechanical replacement is regulated by the adopted mechanical code.", "fact": {"statement": "Replacing ground-level exterior commercial HVAC equipment.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E2"]}, "permit": "VERIFIED_REQUIRED", "permit_finding": "Mechanical permit requirement is established by the cited permit/code evidence.", "permit_basis": "DIRECT_EVIDENCE", "permit_evidence": ["E2"], "pathway": "CONDITIONAL", "pathway_finding": "Specific review pathway remains unresolved because project-specific pathway criteria have not been verified.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Project-specific review pathway criteria (e.g., unit weight, CFM)."], "reopen": ["Authoritative pathway criteria retrieved."]},
                {"type": "Electrical", "applicability": {"rule": "Electrical permit required for branch circuit/disconnect modification.", "fact": {"statement": "Electrical modifications to replacement unit are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Whether wiring/disconnect/breaker will be modified.", "relationship": "conditional", "evidence": ["E4"]}, "permit": "CONDITIONAL", "permit_finding": "Electrical permit consequence depends on whether wiring/disconnect/circuit work occurs.", "permit_basis": "CONDITIONAL", "permit_evidence": ["E4"], "pathway": "CONDITIONAL", "pathway_finding": "Pathway cannot be determined until electrical scope is defined.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Unit electrical specs (MCA, MOP, voltage).", "Scope of electrical changes."], "reopen": ["Modifying electrical disconnect, wiring, or breaker."]},
                {"type": "Energy", "applicability": {"rule": "Current energy code applies to replacement equipment.", "fact": {"statement": "Replacing HVAC unit; efficiency ratings unknown.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E3"]}, "permit": "CONDITIONAL", "permit_finding": "Energy code compliance applies, separate permit unconfirmed.", "permit_basis": "CONDITIONAL", "permit_evidence": [], "pathway": "CONDITIONAL", "pathway_finding": "Compliance pathway depends on equipment specifications.", "pathway_basis": "CONDITIONAL", "pathway_evidence": [], "missing": ["Replacement equipment efficiency/specifications."], "reopen": []},
                {"type": "Structural", "applicability": {"rule": "Structural permits required for alterations or added loads.", "fact": {"statement": "Replacement unit weight and anchorage configuration are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Replacement unit operating weight and anchorage configuration.", "relationship": "not_established", "evidence": []}, "permit": "UNKNOWN", "permit_finding": "Permit consequence cannot be determined without equipment specifications.", "permit_basis": "NOT_ESTABLISHED", "permit_evidence": [], "pathway": "UNKNOWN", "pathway_finding": "Review pathway cannot be established.", "pathway_basis": "NOT_ESTABLISHED", "pathway_evidence": [], "missing": ["Replacement unit operating weight.", "Anchorage configuration."], "reopen": ["Rooftop mounting, suspended installation, or structural framing modifications occur."]},
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

DISCIPLINE DISCOVERY: Determine disciplines dynamically from project scope, project type, jurisdiction, adopted codes, permit requirements, land-use controls, and authoritative applicability rules. Do not use a fixed discipline list. Do not impose a maximum number of disciplines.

For every discipline determine separately:
1. APPLICABILITY: Does the authoritative rule apply to the known project facts?
2. PERMIT: Does authoritative evidence establish a permit requirement or exemption?
3. PATHWAY: Does authoritative evidence establish a specific review pathway?

REGULATORY FIREWALL:
SOURCE RULE + PROJECT FACT does NOT automatically prove a permit or pathway.
Never convert: threshold → permit; threshold → exemption; exemption → pathway; code applicability → permit; permit → plan-review pathway; existing entitlement → exemption.
Permit = VERIFIED_REQUIRED only with permit-specific evidence.
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

EVIDENCE: Every evidence item must support one specific proposition. Evidence discipline labels are descriptive metadata, not proof of applicability. Official building-code provisions may legitimately support multiple related disciplines. Do not reject evidence solely because its discipline label differs from the project discipline; evaluate the actual rule proposition.

CODE CURRENCY: Never assume the current code edition from memory. Verify the edition and effective/mandatory status from the authoritative adoption source for the project date. If code currency is unresolved, use CONDITIONAL.

BOTTOM LINE: Bottom Line may summarize only established findings. It may not introduce a new requirement, exemption, threshold, pathway, CUP conclusion, or code edition. Preserve material uncertainty.

RESEARCH COMPLETENESS: SUFFICIENT = material conclusions supported by adequate authoritative evidence. PARTIAL = main framework established but material facts/documents remain unresolved. INSUFFICIENT = jurisdiction, governing code, permit authority, or material requirements cannot be established.

OUTPUT: Keep JSON concise. Rule: 10-25 words. Fact: 5-15 words. Finding: 10-25 words. Missing/reopen item: short phrase. Research quality is more important than brevity.

JSON SCHEMA:
{{
  "bottom_line": "3 concise sentences maximum",
  "bottom_line_evidence": ["E1"],
  "research_completeness": {{"status": "SUFFICIENT|PARTIAL|INSUFFICIENT", "reason": "short", "critical_missing": ["short item"]}},
  "jurisdiction": {{"status": "VERIFIED|CONDITIONAL", "county": "string", "city": "string", "ahj": "string", "evidence": ["E1"]}},
  "codes": [{{"name": "string", "status": "CURRENT|CONDITIONAL", "evidence": ["E2"]}}],
  "evidence": [{{"id": "E1", "title": "string", "url": "string", "authority": "state|county|city|federal|tribal|other", "discipline": "string", "source_type": "code|ordinance|permit_page|checklist|application|interpretation|entitlement|other", "retrieval_note": "short", "rule": "specific proposition supported"}}],
  "disciplines": [{{
    "type": "string",
    "applicability": {{"rule": "short", "fact": {{"statement": "short", "source": "USER_PROVIDED|RETRIEVED_RECORD|AUTHORITATIVE_SOURCE|INFERRED|UNKNOWN"}}, "determination": "applies|does_not_apply|cannot_determine", "missing": "short", "relationship": "direct|conditional|not_established", "evidence": ["E1"]}},
    "permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "permit_finding": "short",
    "permit_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
    "permit_evidence": ["E2"],
    "pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "pathway_finding": "short",
    "pathway_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
    "pathway_evidence": ["E3"],
    "missing": ["short"],
    "reopen": ["short"]
  }}]
}}
"""
                result = cached_gemini_call(prompt_hash, prompt)

                if result.get("retry"):
                    retry_prompt = f"""
You are completing a regulatory research dossier that previously hit the generation limit.

Return ONLY the required JSON object.

IMPORTANT:
- Do not redo unnecessary research.
- Preserve authoritative findings already established.
- Do not omit a discipline merely to save tokens.
- Do not invent missing facts.
- Do not add explanatory prose outside JSON.
- Keep wording extremely concise.
- Use short strings in every field.
- Evidence rules must remain proposition-specific.
- Applicability, permit requirement, and pathway must remain separate.
- Preserve all material evidence IDs.
- If something cannot be established, use UNKNOWN or CONDITIONAL.
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

JSON SCHEMA:
{{
  "bottom_line": "3 concise sentences maximum",
  "bottom_line_evidence": ["E1"],
  "research_completeness": {{"status": "SUFFICIENT|PARTIAL|INSUFFICIENT", "reason": "short", "critical_missing": ["short item"]}},
  "jurisdiction": {{"status": "VERIFIED|CONDITIONAL", "county": "string", "city": "string", "ahj": "string", "evidence": ["E1"]}},
  "codes": [{{"name": "string", "status": "CURRENT|CONDITIONAL", "evidence": ["E2"]}}],
  "evidence": [{{"id": "E1", "title": "string", "url": "string", "authority": "state|county|city|federal|tribal|other", "discipline": "string", "source_type": "code|ordinance|permit_page|checklist|application|interpretation|entitlement|other", "retrieval_note": "short", "rule": "specific proposition supported"}}],
  "disciplines": [{{
    "type": "string",
    "applicability": {{"rule": "short", "fact": {{"statement": "short", "source": "USER_PROVIDED|RETRIEVED_RECORD|AUTHORITATIVE_SOURCE|INFERRED|UNKNOWN"}}, "determination": "applies|does_not_apply|cannot_determine", "missing": "short", "relationship": "direct|conditional|not_established", "evidence": ["E1"]}},
    "permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "permit_finding": "short",
    "permit_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
    "permit_evidence": ["E2"],
    "pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "pathway_finding": "short",
    "pathway_basis": "DIRECT_EVIDENCE|CONDITIONAL|NOT_ESTABLISHED",
    "pathway_evidence": ["E3"],
    "missing": ["short"],
    "reopen": ["short"]
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
