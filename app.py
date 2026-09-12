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

st.set_page_config(page_title="AHJ Research Assistant v24", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v24_compact_retry_reasoning"

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

        if permit == "VERIFIED_REQUIRED" and not item.get("permit_evidence"):
            errors.append(f"{discipline}: VERIFIED_REQUIRED permit requires permit-specific evidence.")
        if pathway == "VERIFIED_REQUIRED" and not item.get("pathway_evidence"):
            errors.append(f"{discipline}: VERIFIED_REQUIRED pathway requires pathway-specific evidence.")

        if relationship == "direct" and not app.get("rule"):
            errors.append(f"{discipline}: direct relationship requires a source rule.")
        if determination == "cannot_determine" and not app.get("missing"):
            errors.append(f"{discipline}: cannot_determine requires missing facts.")
        if app.get("missing") and determination == "does_not_apply":
            errors.append(f"{discipline}: unresolved facts present, but determination is does_not_apply.")
        if permit == "NOT_CURRENTLY_TRIGGERED" and determination == "does_not_apply" and app.get("missing"):
            errors.append(f"{discipline}: NOT_CURRENTLY_TRIGGERED cannot be paired with does_not_apply when missing facts exist.")
        if permit == "NOT_APPLICABLE":
            if determination != "does_not_apply": errors.append(f"{discipline}: NOT_APPLICABLE requires determination=does_not_apply.")
            if not app.get("evidence"): errors.append(f"{discipline}: NOT_APPLICABLE requires authoritative applicability evidence.")

        if relationship == "not_established":
            finding = (item.get("permit_finding") or "").lower()
            dangerous_phrases = ["is required", "requires", "shall require", "automatically triggers", "therefore requires", "must obtain"]
            for phrase in dangerous_phrases:
                if phrase in finding:
                    errors.append(f"{discipline}: relationship is not_established but permit_finding claims downstream requirement.")
                    break

        finding = (item.get("permit_finding") or "").lower()
        threshold_markers = ["lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "feet", "foot", "percent", "%", "section", "chapter", "threshold"]
        if any(marker in finding for marker in threshold_markers) and not (item.get("permit_evidence") or item.get("pathway_evidence")):
            errors.append(f"{discipline}: finding contains threshold/code claim but has no permit/pathway evidence.")
        if pathway == "VERIFIED_REQUIRED" and relationship == "not_established":
            errors.append(f"{discipline}: pathway cannot be verified when relationship is not_established.")

    return errors

def validate_bottom_line(data):
    errors = []
    bottom_line = (data.get("bottom_line") or "").lower()
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
# CACHING & API CALL WITH RETRY
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text, attempt=1):
    time.sleep(1.0)
    debug_info = {"status": "processing", "attempt": attempt}
    
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())],
            max_output_tokens=8192, # Kept at 8192 to allow full reasoning depth
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
                # AUTOMATIC COMPACT RETRY
                if attempt == 1:
                    retry_prompt = prompt_text + """

IMPORTANT RETRY INSTRUCTION:
The previous generation exceeded the available output budget. Perform the same research and reasoning, but produce a more compact final dossier.
Do NOT remove disciplines or material evidence. Do NOT collapse applicability, permit, and pathway into one field.
Instead: use concise phrases, eliminate repetition, keep missing facts/reopen conditions concise, and keep the Bottom Line to 3-5 sentences maximum.
Correctness and evidence take priority over stylistic brevity.
"""
                    # Recursively call with attempt=2. The cache key changes, forcing a new API call.
                    return cached_gemini_call(prompt_hash, retry_prompt, attempt=2)
                else:
                    return {"data": None, "error": True, "msg": "Model reached output limit even after compact retry.", "debug": debug_info}
                    
            if finish_reason and finish_reason != "FinishReason.STOP":
                debug_info["error_type"] = "Early Stop"
                return {"data": None, "error": True, "msg": f"API stopped early: {finish_reason}", "debug": debug_info}
        else:
            debug_info["candidates"] = None
            debug_info["prompt_feedback"] = str(getattr(response, "prompt_feedback", None))
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "msg": "No candidates returned.", "debug": debug_info}

        text = getattr(response, "text", None)
        if not text:
            try: text = response.candidates[0].content.parts[0].text
            except Exception: pass
        if not text:
            debug_info["raw_text"] = ""
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "msg": "Empty response.", "debug": debug_info}

        data = None
        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:500]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "error": True, "msg": "Failed to parse JSON.", "debug": debug_info}

        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            
        debug_info["status"] = "success"
        return {"data": data, "error": False, "debug": debug_info}
        
    except Exception as e:
        error_msg = str(e)
        debug_info["exception"] = error_msg
        debug_info["error_type"] = "Python Exception"
        if "429" in error_msg:
            return {"data": None, "error": True, "msg": "Quota exceeded.", "debug": debug_info}
        return {"data": None, "error": True, "msg": f"Error: {error_msg[:200]}", "debug": debug_info}

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state: st.session_state.error_msg = None

st.title("🏛️ AHJ Research Assistant v24")
st.caption("Automatic compact retry. Fact provenance. Separated evidence. Traceable synthesis.")

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
    existing_permit = st.text_input("Existing Entitlements (Optional)", "e.g., Existing CUP")

st.header("2. Scope of Work (SOW)")
sow_text = st.text_area("Paste the complete Scope of Work below.", height=200,
    value="Ground-level exterior HVAC unit replacement (like-for-like replacement with a different brand on an existing exterior pad; parcel operates under an existing Conditional Use Permit). should be in exact same place, ductwork wont be affected, no roof penetrations")

st.header("3. Research Execution")
input_string = f"{PROMPT_VERSION}|{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    st.session_state.error_msg = None
    
    if mock_mode:
        st.session_state.report_data = {
            "bottom_line": "Mechanical permit required. Energy code applies, compliance unknown. Electrical scope unknown (CONDITIONAL). Structural/Planning cannot be determined without missing facts.",
            "bottom_line_evidence": ["E2", "E3", "E4"],
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Jurisdiction", "rule": "Jurisdiction for unincorporated areas requires county confirmation."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Mechanical", "rule": "Commercial mechanical permit required for HVAC replacement."},
                {"id": "E3", "title": "Oregon Energy Efficiency Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Energy", "rule": "Replacement mechanical equipment must comply with current energy efficiency standards."},
                {"id": "E4", "title": "Oregon Electrical Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Electrical", "rule": "Electrical permit required for modification of branch circuits or disconnects."}
            ],
            "disciplines": [
                {
                    "type": "Mechanical",
                    "applicability": {"rule": "Commercial mechanical permit required for HVAC replacement.", "fact": {"statement": "Replacing ground-level exterior HVAC unit on existing pad.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E2"]},
                    "permit": "VERIFIED_REQUIRED", "permit_finding": "Commercial mechanical permit required.", "permit_evidence": ["E2"],
                    "pathway": "CONDITIONAL", "pathway_finding": "Pathway depends on unit weight/CFM, not yet established.", "pathway_evidence": [],
                    "missing": ["Proposed unit weight and CFM."], "reopen": []
                },
                {
                    "type": "Electrical",
                    "applicability": {"rule": "Electrical permit required for branch circuit/disconnect modification.", "fact": {"statement": "Electrical modifications to replacement unit are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Whether wiring/disconnect/breaker will be modified.", "relationship": "conditional", "evidence": ["E4"]},
                    "permit": "CONDITIONAL", "permit_finding": "Permit required only if electrical work is modified.", "permit_evidence": ["E4"],
                    "pathway": "CONDITIONAL", "pathway_finding": "Pathway cannot be determined until electrical scope is defined.", "pathway_evidence": [],
                    "missing": ["Unit electrical specs (MCA, MOP, voltage).", "Scope of electrical changes."], "reopen": ["Modifying electrical disconnect, wiring, or breaker."]
                },
                {
                    "type": "Energy",
                    "applicability": {"rule": "Current energy code applies to replacement equipment.", "fact": {"statement": "Replacing HVAC unit; efficiency ratings unknown.", "source": "USER_PROVIDED"}, "determination": "applies", "missing": "", "relationship": "direct", "evidence": ["E3"]},
                    "permit": "CONDITIONAL", "permit_finding": "Energy code compliance applies, separate permit unconfirmed.", "permit_evidence": [],
                    "pathway": "CONDITIONAL", "pathway_finding": "Compliance pathway depends on equipment specifications.", "pathway_evidence": [],
                    "missing": ["Replacement equipment efficiency/specifications."], "reopen": []
                },
                {
                    "type": "Structural",
                    "applicability": {"rule": "Structural permits required for alterations or added loads.", "fact": {"statement": "Replacement unit weight and anchorage configuration are unknown.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Replacement unit operating weight and anchorage configuration.", "relationship": "not_established", "evidence": []},
                    "permit": "NOT_CURRENTLY_TRIGGERED", "permit_finding": "No structural alteration currently established.", "permit_evidence": [],
                    "pathway": "NOT_CURRENTLY_TRIGGERED", "pathway_finding": "No structural review pathway currently triggered.", "pathway_evidence": [],
                    "missing": ["Replacement unit operating weight.", "Anchorage configuration."], "reopen": ["Rooftop mounting, suspended installation, or structural framing modifications occur."]
                },
                {
                    "type": "Planning / CUP",
                    "applicability": {"rule": "Work must comply with existing CUP conditions.", "fact": {"statement": "Parcel operates under existing CUP; actual conditions not retrieved.", "source": "USER_PROVIDED"}, "determination": "cannot_determine", "missing": "Actual CUP conditions governing exterior equipment.", "relationship": "conditional", "evidence": ["E1"]},
                    "permit": "NOT_CURRENTLY_TRIGGERED", "permit_finding": "No current evidence establishes CUP modification is required.", "permit_evidence": [],
                    "pathway": "CONDITIONAL", "pathway_finding": "Final land-use determination conditional on review of existing CUP conditions.", "pathway_evidence": [],
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

OUTPUT COMPRESSION RULE:
Keep the final JSON concise and information-dense. Avoid explanatory prose, repetition, introductions, or conclusions that duplicate the matrix.
Do not repeat the same rule in applicability, permit finding, pathway finding, and bottom line.
Target lengths (flexible, prioritize correctness): rule: 10-30 words, project fact: 5-20 words, permit/pathway finding: 10-30 words, missing/reopen items: 3-15 words.

DISCIPLINE REASONING (Keep separate):
1. APPLICABILITY: Does the rule apply?
2. PERMIT: Does evidence establish a permit requirement?
3. PATHWAY: Does evidence establish a specific review pathway?

PROJECT FACT PROVENANCE:
- USER_PROVIDED: Explicitly stated by user.
- RETRIEVED_RECORD: From retrieved project document.
- AUTHORITATIVE_SOURCE: Established by regulatory source.
- INFERRED: Reasonable inference, not explicitly established.
- UNKNOWN: Not established.
Never represent INFERRED/UNKNOWN as definitely true. Do not assume replacement implies new wiring, structural anchorage, etc.

ABSENCE-OF-EVIDENCE: Absence of evidence is not evidence of non-applicability. Use NOT_CURRENTLY_TRIGGERED or CANNOT_DETERMINE if facts are missing.

EVIDENCE CHAIN: Never infer downstream consequences (e.g., threshold -> structural engineering) without explicit authoritative evidence.

DISCIPLINE INDEPENDENCE: Evidence for Discipline A cannot establish requirements for Discipline B without explicit cross-discipline authority.

BOTTOM LINE: Must be traceable to discipline findings. Do not introduce new thresholds, requirements, or definitive conclusions if underlying disciplines are conditional.

JSON SCHEMA:
{{
  "bottom_line": "3-5 sentences maximum.",
  "bottom_line_evidence": ["E1"],
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
    "permit_evidence": ["E2"],
    "pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
    "pathway_finding": "string",
    "pathway_evidence": ["E3"],
    "missing": ["string"],
    "reopen": ["string"]
  }}]
}}
"""
                result = cached_gemini_call(prompt_hash, prompt, attempt=1)
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
        st.markdown(f"- **{code.get('name', 'Unknown')}** ({code.get('status', 'N/A')})")

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
            st.markdown(f"**Pathway:** {item.get('pathway', 'N/A')}")
            st.markdown(f"**Pathway Finding:** {item.get('pathway_finding', 'N/A')}")
            
            st.divider()
            app_ev, permit_ev, pathway_ev = app.get("evidence", []), item.get("permit_evidence", []), item.get("pathway_evidence", [])
            
            if app_ev:
                st.write("**Applicability Evidence:**")
                for eid in app_ev:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})** `[{ev.get('authority', 'other').upper()}]` `[{ev.get('discipline', 'general').upper()}]`")
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
            doc.add_paragraph(f"Permit Finding: {item.get('permit_finding')}")
            doc.add_paragraph(f"Pathway Finding: {item.get('pathway_finding')}")
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
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Dossier.json", mime="application/json", use_container_width=True)import os
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

st.set_page_config(page_title="AHJ Research Assistant v22", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v22_reasoning_engine"

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
    allowed_statuses = {
        "VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED", "UNKNOWN",
        "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED",
    }
    allowed_determinations = {"applies", "does_not_apply", "cannot_determine"}
    allowed_relationships = {"direct", "conditional", "not_established"}

    # Collect evidence
    evidence_items = data.get("evidence", [])
    evidence_by_id = {e.get("id"): e for e in evidence_items if e.get("id")}
    evidence_ids = set(evidence_by_id.keys())

    # Validate evidence records
    for ev in evidence_items:
        eid = ev.get("id", "Unknown")
        if not ev.get("title"): errors.append(f"{eid}: evidence missing title.")
        if not ev.get("url"): errors.append(f"{eid}: evidence missing URL.")
        if not ev.get("authority"): errors.append(f"{eid}: evidence missing authority.")
        if not ev.get("discipline"): errors.append(f"{eid}: evidence missing discipline.")
        if not ev.get("rule"): errors.append(f"{eid}: evidence missing rule (specific proposition).")

    # Validate jurisdiction
    jurisdiction = data.get("jurisdiction", {})
    for eid in jurisdiction.get("evidence", []):
        if eid not in evidence_ids:
            errors.append(f"Jurisdiction references nonexistent evidence ID: {eid}")
    if jurisdiction.get("status") == "VERIFIED" and not jurisdiction.get("evidence"):
        errors.append("Verified jurisdiction must have evidence.")

    # Validate codes
    for code in data.get("codes", []):
        for eid in code.get("evidence", []):
            if eid not in evidence_ids:
                errors.append(f"Code '{code.get('name')}' references nonexistent evidence ID: {eid}")
        if code.get("status") == "CURRENT" and not code.get("evidence"):
            errors.append(f"Current code '{code.get('name')}' must have evidence.")

    # Validate bottom line traceability
    for eid in data.get("bottom_line_evidence", []):
        if eid not in evidence_ids:
            errors.append(f"Bottom Line references nonexistent evidence ID: {eid}")
    if not data.get("bottom_line_evidence"):
        errors.append("Bottom Line must reference supporting evidence IDs.")

    # Validate disciplines
    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        app = item.get("applicability", {})
        determination = app.get("determination")
        relationship = app.get("relationship")
        fact = app.get("fact", {})
        fact_source = fact.get("source") if isinstance(fact, dict) else None
        fact_statement = fact.get("statement", "") if isinstance(fact, dict) else str(fact)

        # Basic enum checks
        if permit not in allowed_statuses: errors.append(f"{discipline}: invalid permit '{permit}'")
        if pathway not in allowed_statuses: errors.append(f"{discipline}: invalid pathway '{pathway}'")
        if determination not in allowed_determinations: errors.append(f"{discipline}: invalid determination '{determination}'")
        if relationship not in allowed_relationships: errors.append(f"{discipline}: invalid relationship '{relationship}'")

        # Fact provenance validation
        if fact_source not in FACT_SOURCES:
            errors.append(f"{discipline}: invalid fact source '{fact_source}'")
        if not fact_statement:
            errors.append(f"{discipline}: applicability fact is missing.")

        # Inferred/unknown facts cannot support definitive conclusions
        if fact_source in {"INFERRED", "UNKNOWN"} and determination == "does_not_apply":
            errors.append(f"{discipline}: cannot establish does_not_apply from inferred/unknown fact.")
        if fact_source == "INFERRED" and relationship == "direct":
            errors.append(f"{discipline}: inferred fact cannot support direct relationship.")

        # Separate evidence validation
        for eid in app.get("evidence", []):
            if eid not in evidence_ids:
                errors.append(f"{discipline}: applicability references nonexistent evidence ID '{eid}'")
        for eid in item.get("permit_evidence", []):
            if eid not in evidence_ids:
                errors.append(f"{discipline}: permit references nonexistent evidence ID '{eid}'")
        for eid in item.get("pathway_evidence", []):
            if eid not in evidence_ids:
                errors.append(f"{discipline}: pathway references nonexistent evidence ID '{eid}'")

        # Cross-discipline evidence protection
        for eid in app.get("evidence", []):
            evidence = evidence_by_id.get(eid)
            if evidence:
                ev_disc = (evidence.get("discipline") or "").lower()
                if ev_disc and discipline.lower() and ev_disc != discipline.lower() and relationship == "direct":
                    errors.append(f"{discipline}: direct applicability relies on {ev_disc} evidence {eid}. Explicit cross-discipline authority required.")

        # VERIFIED_REQUIRED requires specific evidence
        if permit == "VERIFIED_REQUIRED" and not item.get("permit_evidence"):
            errors.append(f"{discipline}: VERIFIED_REQUIRED permit requires permit-specific evidence.")
        if pathway == "VERIFIED_REQUIRED" and not item.get("pathway_evidence"):
            errors.append(f"{discipline}: VERIFIED_REQUIRED pathway requires pathway-specific evidence.")

        # Direct relationship requires an actual rule
        if relationship == "direct" and not app.get("rule"):
            errors.append(f"{discipline}: direct relationship requires a source rule.")

        # cannot_determine requires missing information
        if determination == "cannot_determine" and not app.get("missing"):
            errors.append(f"{discipline}: cannot_determine requires missing facts.")

        # Missing facts cannot coexist with definitive does_not_apply
        if app.get("missing") and determination == "does_not_apply":
            errors.append(f"{discipline}: unresolved facts present, but determination is does_not_apply.")

        # NOT_CURRENTLY_TRIGGERED should not have definitive does_not_apply when unresolved facts remain
        if permit == "NOT_CURRENTLY_TRIGGERED" and determination == "does_not_apply" and app.get("missing"):
            errors.append(f"{discipline}: NOT_CURRENTLY_TRIGGERED cannot be paired with does_not_apply when missing facts could change applicability.")

        # NOT_APPLICABLE requires defensible determination + evidence
        if permit == "NOT_APPLICABLE":
            if determination != "does_not_apply":
                errors.append(f"{discipline}: NOT_APPLICABLE requires determination=does_not_apply.")
            if not app.get("evidence"):
                errors.append(f"{discipline}: NOT_APPLICABLE requires authoritative applicability evidence.")

        # not_established relationship cannot claim verified downstream requirement
        if relationship == "not_established":
            finding = (item.get("permit_finding") or "").lower()
            dangerous_phrases = ["is required", "requires", "shall require", "automatically triggers", "therefore requires", "must obtain"]
            for phrase in dangerous_phrases:
                if phrase in finding:
                    errors.append(f"{discipline}: relationship is not_established but permit_finding claims downstream requirement: '{phrase}'.")
                    break

        # Specific threshold/pathway claims should have evidence
        finding = (item.get("permit_finding") or "").lower()
        threshold_markers = ["lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "feet", "foot", "percent", "%", "section", "chapter", "threshold"]
        if any(marker in finding for marker in threshold_markers) and not (item.get("permit_evidence") or item.get("pathway_evidence")):
            errors.append(f"{discipline}: finding contains threshold/code claim but has no permit/pathway evidence.")

        # Pathway cannot be more certain than the evidence
        if pathway == "VERIFIED_REQUIRED" and relationship == "not_established":
            errors.append(f"{discipline}: pathway cannot be verified when relationship is not_established.")

    return errors

def validate_bottom_line(data):
    """Check that bottom line doesn't overstate conditional findings."""
    errors = []
    bottom_line = (data.get("bottom_line") or "").lower()

    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        app = item.get("applicability", {})
        determination = app.get("determination")

        unresolved = (
            permit in {"CONDITIONAL", "UNKNOWN", "NOT_CURRENTLY_TRIGGERED"}
            or pathway in {"CONDITIONAL", "UNKNOWN", "NOT_CURRENTLY_TRIGGERED"}
            or determination == "cannot_determine"
        )

        if not unresolved:
            continue

        discipline_words = [discipline.lower(), discipline.lower().replace("/", " ")]
        if not any(word in bottom_line for word in discipline_words if word):
            continue

        risky_patterns = [
            f"{discipline.lower()} permit is required",
            f"{discipline.lower()} permit required",
            f"{discipline.lower()} approval is required",
            f"{discipline.lower()} approval required",
            f"{discipline.lower()} is not currently triggered",
            f"{discipline.lower()} not currently triggered",
        ]

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
    debug_info = {"status": "processing"}
    
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
                return {"data": None, "error": True, "msg": "Model reached the output limit during research. The request may require a second research pass.", "debug": debug_info}
                
            if finish_reason and finish_reason != "FinishReason.STOP":
                debug_info["error_type"] = "Early Stop"
                return {"data": None, "error": True, "msg": f"API stopped early: {finish_reason}", "debug": debug_info}
        else:
            debug_info["candidates"] = None
            debug_info["prompt_feedback"] = str(getattr(response, "prompt_feedback", None))
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "error": True, "msg": "No candidates returned.", "debug": debug_info}

        text = getattr(response, "text", None)
        if not text:
            try: text = response.candidates[0].content.parts[0].text
            except Exception: pass
                
        if not text:
            debug_info["raw_text"] = ""
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "error": True, "msg": "Empty response.", "debug": debug_info}

        data = None
        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:500]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "error": True, "msg": "Failed to parse JSON.", "debug": debug_info}

        validation_errors = validate_dossier(data)
        validation_errors.extend(validate_bottom_line(data))
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            
        debug_info["status"] = "success"
        return {"data": data, "error": False, "debug": debug_info}
        
    except Exception as e:
        error_msg = str(e)
        debug_info["exception"] = error_msg
        debug_info["error_type"] = "Python Exception"
        if "429" in error_msg:
            return {"data": None, "error": True, "msg": "Quota exceeded.", "debug": debug_info}
        return {"data": None, "error": True, "msg": f"Error: {error_msg[:200]}", "debug": debug_info}

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}
if "error_msg" not in st.session_state: st.session_state.error_msg = None

st.title("🏛️ AHJ Research Assistant v22")
st.caption("Reasoning engine with fact provenance, separated evidence, and traceable synthesis.")

with st.sidebar:
    st.warning("⚠️ Pay-As-You-Go Active. Results cached for 1 hour.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=False)

# --- SECTION 1: PROJECT METADATA ---
st.header("1. Project Metadata")
col1, col2 = st.columns(2)
with col1:
    state = st.selectbox("State / Jurisdiction", STATE_OPTIONS, index=STATE_OPTIONS.index("Oregon"))
    address = st.text_input("Project Address", "24340 NW Meek Rd, 97124")
    project_date = st.date_input("Permit / Construction Date", date.today())
with col2:
    ptype = st.selectbox("Project Type", PROJECT_TYPES, index=0)
    bclass = st.selectbox("Building / Occupancy Class", BUILDING_CLASSES, index=0)
    existing_permit = st.text_input("Existing Entitlements (Optional)", "e.g., Existing CUP")

# --- SECTION 2: SCOPE OF WORK ---
st.header("2. Scope of Work (SOW)")
sow_text = st.text_area(
    "Paste the complete Scope of Work below.", 
    height=200,
    value="Ground-level exterior HVAC unit replacement (like-for-like replacement with a different brand on an existing exterior pad; parcel operates under an existing Conditional Use Permit). should be in exact same place, ductwork wont be affected, no roof penetrations"
)

# --- SECTION 3: RESEARCH ---
st.header("3. Research Execution")
input_string = f"{PROMPT_VERSION}|{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    st.session_state.error_msg = None
    
    if mock_mode:
        st.session_state.report_data = {
            "bottom_line": "Mechanical permit required (VERIFIED_REQUIRED). Energy code applies but compliance unknown. Electrical scope unknown — permit CONDITIONAL. Structural and Planning cannot be determined without missing facts.",
            "bottom_line_evidence": ["E2", "E3", "E4"],
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Jurisdiction", "rule": "Jurisdiction for unincorporated areas requires county confirmation."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Mechanical", "rule": "Commercial mechanical permit required for HVAC replacement."},
                {"id": "E3", "title": "Oregon Energy Efficiency Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Energy", "rule": "Replacement mechanical equipment must comply with current energy efficiency standards."},
                {"id": "E4", "title": "Oregon Electrical Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Electrical", "rule": "Electrical permit required for modification of branch circuits or disconnects."}
            ],
            "disciplines": [
                {
                    "type": "Mechanical",
                    "applicability": {
                        "rule": "Commercial mechanical permit required for HVAC replacement.",
                        "fact": {"statement": "Replacing ground-level exterior HVAC unit with different brand on existing pad.", "source": "USER_PROVIDED"},
                        "determination": "applies",
                        "missing": "",
                        "relationship": "direct",
                        "evidence": ["E2"]
                    },
                    "permit": "VERIFIED_REQUIRED",
                    "permit_finding": "Commercial mechanical permit required for replacement.",
                    "permit_evidence": ["E2"],
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Pathway (minor vs. plan review) depends on unit weight/CFM, not yet established.",
                    "pathway_evidence": [],
                    "missing": ["Proposed unit weight and CFM for pathway determination."],
                    "reopen": []
                },
                {
                    "type": "Electrical",
                    "applicability": {
                        "rule": "Electrical permit required for modification of branch circuits or disconnects.",
                        "fact": {"statement": "Electrical modifications to the replacement unit are unknown.", "source": "USER_PROVIDED"},
                        "determination": "cannot_determine",
                        "missing": "Whether electrical wiring, disconnect, or breaker will be modified.",
                        "relationship": "conditional",
                        "evidence": ["E4"]
                    },
                    "permit": "CONDITIONAL",
                    "permit_finding": "Permit required only if electrical work is actually modified.",
                    "permit_evidence": ["E4"],
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Pathway cannot be determined until electrical scope is defined.",
                    "pathway_evidence": [],
                    "missing": ["Unit electrical specs (MCA, MOP, voltage)", "Scope of electrical changes."],
                    "reopen": ["Modifying electrical disconnect, wiring, or breaker."]
                },
                {
                    "type": "Energy",
                    "applicability": {
                        "rule": "Current energy code applies to replacement equipment.",
                        "fact": {"statement": "Replacing HVAC unit; efficiency ratings unknown.", "source": "USER_PROVIDED"},
                        "determination": "applies",
                        "missing": "",
                        "relationship": "direct",
                        "evidence": ["E3"]
                    },
                    "permit": "CONDITIONAL",
                    "permit_finding": "Energy code compliance applies, but separate energy permit requirement is unconfirmed.",
                    "permit_evidence": [],
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Compliance pathway depends on equipment specifications.",
                    "pathway_evidence": [],
                    "missing": ["Replacement equipment efficiency/specifications to prove compliance."],
                    "reopen": []
                },
                {
                    "type": "Structural",
                    "applicability": {
                        "rule": "Structural permits required for alterations or added loads.",
                        "fact": {"statement": "Replacement unit weight and anchorage configuration are unknown.", "source": "USER_PROVIDED"},
                        "determination": "cannot_determine",
                        "missing": "Replacement unit operating weight and anchorage configuration.",
                        "relationship": "not_established",
                        "evidence": []
                    },
                    "permit": "NOT_CURRENTLY_TRIGGERED",
                    "permit_finding": "No structural alteration currently established from available evidence.",
                    "permit_evidence": [],
                    "pathway": "NOT_CURRENTLY_TRIGGERED",
                    "pathway_finding": "No structural review pathway currently triggered.",
                    "pathway_evidence": [],
                    "missing": ["Replacement unit operating weight", "Anchorage configuration."],
                    "reopen": ["Rooftop mounting, suspended installation, or structural framing modifications occur."]
                },
                {
                    "type": "Planning / CUP",
                    "applicability": {
                        "rule": "Work must comply with existing CUP conditions.",
                        "fact": {"statement": "Parcel operates under existing CUP; actual conditions not retrieved.", "source": "USER_PROVIDED"},
                        "determination": "cannot_determine",
                        "missing": "Actual CUP conditions governing exterior equipment.",
                        "relationship": "conditional",
                        "evidence": ["E1"]
                    },
                    "permit": "NOT_CURRENTLY_TRIGGERED",
                    "permit_finding": "No current evidence establishes a CUP modification is required for like-for-like replacement.",
                    "permit_evidence": [],
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Final land-use determination conditional on review of existing CUP conditions.",
                    "pathway_evidence": [],
                    "missing": ["Actual CUP conditions governing exterior equipment."],
                    "reopen": ["Relocation, footprint expansion, screening changes, noise increases, or site work occur."]
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
You are an expert AHJ research analyst. Return ONLY a valid JSON object. Do not include any conversational text, markdown formatting, or explanations outside the JSON.

PROJECT:
State: {state} | Address: {address} | Date: {project_date}
Type: {ptype} | Class: {bclass} (USER-PROVIDED) | Entitlements: {existing_permit}
SCOPE: {sow_text}

RESEARCH OBJECTIVE:
Determine jurisdiction, current applicable code editions, permit requirements, review pathways, conditional triggers, and unresolved facts.

DISCIPLINE REASONING FRAMEWORK
For every potentially relevant discipline, reason through these questions independently:
1. APPLICABILITY: Does an authoritative rule apply to the project?
2. PERMIT: Does authoritative evidence establish that the rule creates a permit requirement?
3. PATHWAY: Does authoritative evidence establish a specific review, submittal, inspection, or approval pathway?
4. MISSING FACTS: What facts or documents are still needed?
5. REOPEN CONDITIONS: What future or newly discovered facts would cause this discipline to become relevant or change its conclusion?
Do not collapse them into one conclusion.

PROJECT FACT PROVENANCE RULE
Every project fact must be classified by source:
- USER_PROVIDED: Explicitly stated by the user/project input.
- RETRIEVED_RECORD: Established by a retrieved project document, permit record, approval, drawing, or other project-specific record.
- AUTHORITATIVE_SOURCE: Established directly by an authoritative regulatory source.
- INFERRED: A reasonable inference from known facts, but not explicitly established.
- UNKNOWN: Not established.

Never represent an INFERRED or UNKNOWN condition as a project fact that is definitely true.
Do not assume that equipment replacement means: new wiring, new disconnect, breaker modification, structural anchorage, structural engineering, planning approval, fire review, utility work, or any other downstream scope. Those must be established independently.

ABSENCE-OF-EVIDENCE RULE
Never conclude that a requirement does not apply merely because research did not find evidence establishing the requirement.
Distinguish:
1. NOT_APPLICABLE: Authoritative evidence establishes that the rule does not apply to this project.
2. NOT_CURRENTLY_TRIGGERED: No current trigger has been established, but additional facts or documents could change the result.
3. CANNOT_DETERMINE: Available facts or governing documents are insufficient to determine applicability.
4. UNKNOWN: The relevant rule or requirement has not been sufficiently established.
Absence of evidence is not evidence of non-applicability.

CONDITIONAL GOVERNING-DOCUMENT RULE
If a project is governed by an existing approval, permit, CUP, variance, development agreement, site plan, or similar document, do not conclude that the approval is unaffected unless the actual governing conditions have been retrieved and reviewed. If the project appears like-for-like but the governing conditions are unavailable: applicability="cannot_determine", relationship="conditional".

APPLICABILITY VS COMPLIANCE
Do not confuse whether a regulation applies with whether the project has demonstrated compliance.
- The energy code may APPLY.
- Equipment efficiency information may be UNKNOWN.
- A separate energy permit may NOT be established.
Therefore these may legitimately coexist: applicability="applies", permit="CONDITIONAL", missing="equipment compliance information".

UNRESOLVED FACT RULE
If an unresolved project fact could change whether a regulation applies, use determination="cannot_determine". Do not use "does_not_apply" when the missing fact could change the conclusion. However, if the rule's applicability is already established and the missing information only affects compliance, permit pathway, or documentation, the applicability determination may remain "applies" while permit/pathway remain CONDITIONAL or UNKNOWN.

EVIDENCE CHAIN RULE
Never infer a downstream regulatory consequence from a threshold, exemption, definition, scope rule, or general requirement unless authoritative evidence explicitly establishes that consequence.
Example: If a source establishes "Equipment over X is outside exemption Y", you may conclude "Exemption Y does not apply if the equipment exceeds X." You may NOT automatically conclude "Structural engineering is required" unless authoritative structural evidence establishes that relationship.

DISCIPLINE INDEPENDENCE RULE
Evidence discovered for Discipline A cannot establish a regulatory requirement for Discipline B unless the authoritative source explicitly connects the two. When a cross-discipline consequence is suspected but not explicitly established, mark the relationship "not_established" and explain what authoritative evidence would be required.

EVIDENCE STRENGTH RULE
A source can establish only the proposition it actually supports. Do not expand: definition → requirement, threshold → consequence, exemption → opposite requirement, permit → review pathway, code applicability → permit requirement, project fact → compliance, absence of evidence → non-applicability. Each transition requires either explicit authoritative evidence or a clearly labeled conditional inference.

BOTTOM LINE TRACEABILITY
Every material statement in the Bottom Line must be supported by the underlying discipline findings and cited evidence (via bottom_line_evidence). The Bottom Line may summarize or combine findings. It may NOT introduce: a new threshold, a new permit requirement, a new exemption, a new pathway, a new jurisdiction, or a new definitive conclusion. If a material discipline is conditional or unresolved, preserve that uncertainty in the Bottom Line.

MISSING INFORMATION CLASSIFICATION
Do not use "missing" as a generic dumping ground. Classify missing information according to what it affects: applicability, compliance, permit, or pathway. A missing compliance document does not necessarily make applicability uncertain. A missing governing condition may make applicability uncertain. A missing equipment specification may affect pathway without affecting permit applicability.

CRITICAL RULES:
1. NO ARTIFICIAL CAPS: Do not limit the number of disciplines or evidence items.
2. JURISDICTION: Establish the actual building/AHJ jurisdiction. If unclear, set status="CONDITIONAL" and ahj="Unresolved - boundary unconfirmed".
3. FEDERAL/TRIBAL/HISTORIC: Actively check for federal waterways, tribal land, or historic districts. Add a specific discipline if applicable.
4. STATUS VALUES: VERIFIED_REQUIRED, CONDITIONAL, INFERRED, UNKNOWN, NOT_APPLICABLE, NOT_CURRENTLY_TRIGGERED, USER_PROVIDED.
5. DETERMINATION VALUES: applies, does_not_apply, cannot_determine.
6. RELATIONSHIP VALUES: direct, conditional, not_established.
7. FACT SOURCE VALUES: USER_PROVIDED, RETRIEVED_RECORD, AUTHORITATIVE_SOURCE, INFERRED, UNKNOWN.
8. DO NOT SPECULATE: Never invent dates, thresholds, or requirements without retrieved evidence.

JSON SCHEMA:
{{
  "bottom_line": "Concise executive summary of the research findings.",
  "bottom_line_evidence": ["E1", "E2"],
  "jurisdiction": {{
    "status": "VERIFIED or CONDITIONAL",
    "county": "string",
    "city": "string",
    "ahj": "string",
    "evidence": ["E1"]
  }},
  "codes": [
    {{"name": "string", "status": "CURRENT or CONDITIONAL", "evidence": ["E2"]}}
  ],
  "evidence": [
    {{
      "id": "E1", 
      "title": "Official source title", 
      "url": "string", 
      "authority": "state|county|city|federal|tribal|other",
      "discipline": "Mechanical|Electrical|Structural|Planning|Energy|Jurisdiction|etc",
      "rule": "The specific regulatory proposition supported by this source."
    }}
  ],
  "disciplines": [
    {{
      "type": "string",
      "applicability": {{
        "rule": "What the authoritative source actually establishes.",
        "fact": {{
          "statement": "What is actually known about this project.",
          "source": "USER_PROVIDED|RETRIEVED_RECORD|AUTHORITATIVE_SOURCE|INFERRED|UNKNOWN"
        }},
        "determination": "applies|does_not_apply|cannot_determine",
        "missing": "What fact prevents a final applicability determination, if any.",
        "relationship": "direct|conditional|not_established",
        "evidence": ["E1"]
      }},
      "permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
      "permit_finding": "What the evidence establishes about the permit requirement.",
      "permit_evidence": ["E2"],
      "pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
      "pathway_finding": "What the evidence establishes about the review/submittal pathway.",
      "pathway_evidence": ["E3"],
      "missing": ["Facts or documents still needed to determine permit or pathway."],
      "reopen": ["Conditions that would cause this discipline to be reconsidered."]
    }}
  ]
}}
"""
                result = cached_gemini_call(prompt_hash, prompt)
                st.session_state.debug_log = result.get("debug", {})
                
                if result["error"]:
                    st.session_state.error_msg = result["msg"]
                else:
                    st.session_state.report_data = result["data"]

# --- DISPLAY ERRORS & DEBUG LOG ---
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
    if bl_ev:
        st.caption(f"Bottom Line Evidence IDs: {', '.join(bl_ev)}")

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
            if eid in ev_dict:
                st.write(f"**Source:** [{ev_dict[eid]['title']}]({ev_dict[eid]['url']})")

    st.subheader("📚 Applicable Codes")
    for code in (data.get("codes") or []):
        st.markdown(f"- **{code.get('name', 'Unknown')}** ({code.get('status', 'N/A')})")

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
            st.markdown(f"**Pathway:** {item.get('pathway', 'N/A')}")
            st.markdown(f"**Pathway Finding:** {item.get('pathway_finding', 'N/A')}")
            
            st.divider()
            
            # Show separated evidence
            app_ev = app.get("evidence", [])
            permit_ev = item.get("permit_evidence", [])
            pathway_ev = item.get("pathway_evidence", [])
            
            if app_ev:
                st.write("**Applicability Evidence:**")
                for eid in app_ev:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})** `[{ev.get('authority', 'other').upper()}]` `[{ev.get('discipline', 'general').upper()}]`")
                        st.caption(f"  *Rule:* {ev.get('rule', 'N/A')}")
            
            if permit_ev:
                st.write("**Permit Evidence:**")
                for eid in permit_ev:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})**")
            else:
                st.caption("*No permit-specific evidence retrieved.*")
            
            if pathway_ev:
                st.write("**Pathway Evidence:**")
                for eid in pathway_ev:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})**")
            else:
                st.caption("*No pathway-specific evidence retrieved.*")
            
            missing = item.get("missing", [])
            if isinstance(missing, str): missing = [missing]
            if missing:
                st.markdown("**Missing Information:**")
                for fact in missing:
                    st.markdown(f"- {fact}")
                    
            reopen = item.get("reopen", [])
            if isinstance(reopen, str): reopen = [reopen]
            if reopen:
                st.markdown("**Reopen If:**")
                for condition in reopen:
                    st.markdown(f"- {condition}")

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
        for code in (data.get("codes") or []): 
            doc.add_paragraph(f"{code.get('name')} ({code.get('status')})", style='List Bullet')
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
            doc.add_paragraph(f"Permit Finding: {item.get('permit_finding')}")
            doc.add_paragraph(f"Pathway Finding: {item.get('pathway_finding')}")
            
            missing = item.get("missing", [])
            if isinstance(missing, str): missing = [missing]
            if missing:
                doc.add_paragraph("Missing Information:", style='List Bullet')
                for fact in missing:
                    doc.add_paragraph(f"  - {fact}")
                    
            reopen = item.get("reopen", [])
            if isinstance(reopen, str): reopen = [reopen]
            if reopen:
                doc.add_paragraph("Reopen If:", style='List Bullet')
                for condition in reopen:
                    doc.add_paragraph(f"  - {condition}")
                
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📄 Download Word Report", data=buf.getvalue(), file_name="AHJ_Dossier.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with col2:
        json_data = json.dumps({"project": {"state": state, "address": address, "date": str(project_date)}, "dossier": data}, indent=2)
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Dossier.json", mime="application/json", use_container_width=True)
