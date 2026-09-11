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

st.set_page_config(page_title="AHJ Research Assistant v18", page_icon="🏛️", layout="wide")

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

GEMINI_KEY = os.getenv("GEMINI_KEY") or st.secrets.get("GEMINI_KEY", "")
PROMPT_VERSION = "v18_strict_evidence_validator"

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
    allowed_permits = {"VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED", "UNKNOWN", "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED"}
    allowed_determinations = {"applies", "does_not_apply", "cannot_determine"}
    
    evidence_ids = {ev["id"] for ev in data.get("evidence", [])}
    
    for item in data.get("disciplines", []):
        permit = item.get("permit")
        pathway = item.get("pathway")
        test = item.get("test", {})
        determination = test.get("determination")
        ev_ids = item.get("evidence", [])
        missing = test.get("missing", "").lower()
        
        if permit not in allowed_permits: errors.append(f"{item.get('type')}: invalid permit '{permit}'")
        if pathway not in allowed_permits: errors.append(f"{item.get('type')}: invalid pathway '{pathway}'")
        if determination and determination not in allowed_determinations: errors.append(f"{item.get('type')}: invalid determination '{determination}'")
        
        # Rule: VERIFIED_REQUIRED must have evidence
        if permit == "VERIFIED_REQUIRED" and not ev_ids:
            errors.append(f"{item.get('type')}: VERIFIED_REQUIRED must have evidence IDs")
            
        # Rule: NOT_APPLICABLE cannot be cannot_determine
        if permit == "NOT_APPLICABLE" and determination == "cannot_determine":
            errors.append(f"{item.get('type')}: NOT_APPLICABLE conflicts with cannot_determine")
            
        # Rule: NOT_CURRENTLY_TRIGGERED with missing facts must be cannot_determine
        if permit == "NOT_CURRENTLY_TRIGGERED" and missing not in ["none", "n/a", ""]:
            if determination != "cannot_determine":
                errors.append(f"{item.get('type')}: NOT_CURRENTLY_TRIGGERED with missing facts must be cannot_determine")
                
        # Rule: Check evidence IDs exist
        for eid in ev_ids:
            if eid not in evidence_ids:
                errors.append(f"{item.get('type')}: references non-existent evidence ID '{eid}'")
                
    # Rule: CURRENT codes must have evidence
    for code in data.get("codes", []):
        if code.get("status") == "CURRENT" and not code.get("evidence"):
            errors.append(f"Code {code.get('name')}: CURRENT status must have evidence IDs")
            
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

st.title("🏛️ AHJ Research Assistant v18")
st.caption("Strict evidence validation. Applicability testing. No artificial caps.")

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
            "bottom_line": "Mechanical permit VERIFIED REQUIRED. Electrical and Energy CONDITIONAL. Jurisdiction needs confirmation.",
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "rule": "Jurisdiction for unincorporated areas."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "rule": "Commercial mechanical permit required for replacement."}
            ],
            "disciplines": [
                {
                    "type": "Mechanical", "permit": "VERIFIED_REQUIRED", "pathway": "CONDITIONAL",
                    "finding": "Commercial mechanical permit required for replacement.",
                    "evidence": ["E2"],
                    "test": {"rule": "Permit required for regulated mechanical replacement.", "fact": "Commercial HVAC replacement.", "determination": "applies", "missing": "Unit specs for pathway determination."},
                    "reopen": ""
                },
                {
                    "type": "Electrical", "permit": "CONDITIONAL", "pathway": "CONDITIONAL",
                    "finding": "Permit required only if electrical work is modified.",
                    "evidence": ["E1"],
                    "test": {"rule": "Permits required for branch circuit/disconnect modification.", "fact": "SOW states different brand, electrical scope unstated.", "determination": "cannot_determine", "missing": "Unit electrical specs and scope of electrical changes."},
                    "reopen": "Modifying electrical disconnect, wiring, or breaker."
                },
                {
                    "type": "Energy", "permit": "CONDITIONAL", "pathway": "CONDITIONAL",
                    "finding": "Replacement equipment must be checked against current energy requirements.",
                    "evidence": ["E2"],
                    "test": {"rule": "Current energy code applies to replacement equipment.", "fact": "Replacing HVAC unit; efficiency ratings unknown.", "determination": "cannot_determine", "missing": "Replacement equipment efficiency/specifications."},
                    "reopen": ""
                },
                {
                    "type": "Planning / CUP", "permit": "NOT_CURRENTLY_TRIGGERED", "pathway": "NOT_CURRENTLY_TRIGGERED",
                    "finding": "No new land-use trigger identified from current scope.",
                    "evidence": ["E1"],
                    "test": {"rule": "Work must comply with existing CUP conditions.", "fact": "Parcel operates under existing CUP; conditions not retrieved.", "determination": "cannot_determine", "missing": "Actual CUP conditions governing exterior equipment."},
                    "reopen": "Relocation, footprint expansion, screening changes, or site work occur."
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

CRITICAL RULES:
1. NO ARTIFICIAL CAPS: Do not limit the number of disciplines or evidence items. If the research requires 10 disciplines and 15 evidence sources, output them all.
2. STRICT EVIDENCE LINKING: A finding may ONLY state a specific threshold, permit requirement, exemption, code section, or review pathway when an evidence item's "rule" explicitly supports that statement. DO NOT infer downstream requirements (e.g., plan review, engineered seismic anchoring) from a threshold alone unless the source explicitly establishes that relationship.
3. ENERGY DISCIPLINE: For every HVAC/mechanical project, ALWAYS evaluate Energy as a separate discipline. Determine compliance requirements from authoritative evidence, do not assume.
4. JURISDICTION: Establish the actual building/AHJ jurisdiction from authoritative evidence. If the boundary is unclear, set status="CONDITIONAL" and ahj="Unresolved - boundary unconfirmed".
5. FEDERAL/TRIBAL/HISTORIC: Actively check for federal waterways, tribal land, or historic districts. If applicable, add a specific discipline.
6. PERMIT VS PATHWAY: Keep separate. If authoritative evidence establishes a permit is required, permit="VERIFIED_REQUIRED" even if pathway="CONDITIONAL".
7. STATUS VALUES: VERIFIED_REQUIRED, CONDITIONAL, INFERRED, UNKNOWN, NOT_APPLICABLE, NOT_CURRENTLY_TRIGGERED, USER_PROVIDED.
8. APPLICABILITY TEST: For EVERY discipline, fill out the "test" object. If a material fact is missing, determination MUST be "cannot_determine". NEVER use NOT_APPLICABLE if a missing fact could change the result.
9. DO NOT SPECULATE: Never invent dates, thresholds, or requirements without retrieved evidence.

JSON SCHEMA:
{{
  "bottom_line": "Concise executive summary of the research findings.",
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
      "rule": "Short statement of what the source actually establishes."
    }}
  ],
  "disciplines": [
    {{
      "type": "string",
      "permit": "VERIFIED_REQUIRED or CONDITIONAL or NOT_CURRENTLY_TRIGGERED",
      "pathway": "CONDITIONAL or NOT_CURRENTLY_TRIGGERED",
      "finding": "Detailed finding based on retrieved evidence.",
      "evidence": ["E2"],
      "test": {{
        "rule": "Specific rule or threshold found in the evidence.",
        "fact": "Specific fact from the SOW or metadata.",
        "determination": "applies, does_not_apply, or cannot_determine",
        "missing": "What specific fact is missing to close the gap."
      }},
      "reopen": "Specific conditions that would trigger this discipline if currently NOT_CURRENTLY_TRIGGERED."
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
            st.write(f"**Finding:** {item.get('finding', 'N/A')}")
            
            ev_ids = item.get("evidence", [])
            if ev_ids:
                st.write("**Evidence:**")
                for eid in ev_ids:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})**")
                        st.caption(f"  *Rule:* {ev.get('rule', 'N/A')}")
            else:
                st.write("**Evidence:** None retrieved.")
            
            st.divider()
            st.subheader("Applicability Test")
            app_test = item.get("test") or {}
            st.write(f"**Source Rule:** {app_test.get('rule', 'N/A')}")
            st.write(f"**Project Fact:** {app_test.get('fact', 'N/A')}")
            st.write(f"**Determination:** {app_test.get('determination', 'N/A').replace('_', ' ').title()}")
            if app_test.get('missing'):
                st.info(f"**❓ Missing:** {app_test.get('missing')}")
            
            reopen = item.get('reopen', '')
            if reopen:
                st.success(f"**🔄 Reopen if:** {reopen}")

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
            doc.add_paragraph(f"Finding: {item.get('finding')}")
            
            app_test = item.get("test") or {}
            doc.add_paragraph(f"Source Rule: {app_test.get('rule')}")
            doc.add_paragraph(f"Project Fact: {app_test.get('fact')}")
            doc.add_paragraph(f"Determination: {app_test.get('determination', '').replace('_', ' ').title()}")
            if app_test.get('missing'):
                doc.add_paragraph(f"Missing: {app_test.get('missing')}")
            if item.get('reopen'):
                doc.add_paragraph(f"Reopen if: {item.get('reopen')}")
                
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📄 Download Word Report", data=buf.getvalue(), file_name="AHJ_Dossier.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with col2:
        json_data = json.dumps({"project": {"state": state, "address": address, "date": str(project_date)}, "dossier": data}, indent=2)
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Dossier.json", mime="application/json", use_container_width=True)
