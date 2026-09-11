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

st.set_page_config(page_title="AHJ Research Assistant v9", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v9_compact_evidence" # Invalidates cache when schema changes

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
        pass
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("No JSON object found")
    return json.loads(text[start:end + 1])

def validate_dossier(data):
    errors = []
    allowed_permits = {"VERIFIED_REQUIRED", "CONDITIONAL", "INFERRED", "UNKNOWN", "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED"}
    allowed_determinations = {"applies", "does_not_apply", "cannot_determine"}
    
    for item in data.get("disciplines", []):
        permit = item.get("permit")
        pathway = item.get("pathway")
        if permit not in allowed_permits:
            errors.append(f"{item.get('type')}: invalid permit status '{permit}'")
        if pathway not in allowed_permits:
            errors.append(f"{item.get('type')}: invalid pathway status '{pathway}'")
        
        test = item.get("test", {})
        if test.get("determination") not in allowed_determinations:
            errors.append(f"{item.get('type')}: invalid determination '{test.get('determination')}'")
            
        if permit == "NOT_APPLICABLE" and test.get("determination") == "cannot_determine":
            errors.append(f"{item.get('type')}: NOT_APPLICABLE conflicts with cannot_determine")
            
    return errors

# ============================================================
# CACHING & API CALL
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0) # RPM throttle
    debug_info = {"status": "processing"}
    
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config = types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())],
            max_output_tokens=5000, # Optimized for compact schema
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
                return {"data": None, "sources": [], "error": True, "msg": "Model hit output limit. Try a shorter SOW.", "debug": debug_info}
                
            if finish_reason and finish_reason != "FinishReason.STOP":
                debug_info["error_type"] = "Early Stop"
                return {"data": None, "sources": [], "error": True, "msg": f"API stopped early: {finish_reason}", "debug": debug_info}
        else:
            debug_info["candidates"] = None
            debug_info["prompt_feedback"] = str(getattr(response, "prompt_feedback", None))
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "sources": [], "error": True, "msg": "No candidates returned.", "debug": debug_info}

        text = getattr(response, "text", None)
        if not text:
            try:
                text = response.candidates[0].content.parts[0].text
            except Exception:
                pass
                
        if not text:
            debug_info["raw_text"] = ""
            debug_info["error_type"] = "Empty Text"
            return {"data": None, "sources": [], "error": True, "msg": "Empty response.", "debug": debug_info}

        data = None
        try:
            data = extract_json(text)
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:500]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "sources": [], "error": True, "msg": "Failed to parse JSON.", "debug": debug_info}

        # Validate the schema
        validation_errors = validate_dossier(data)
        if validation_errors:
            debug_info["validation_errors"] = validation_errors
            # We display warnings but don't hard-fail during tuning
            
        debug_info["status"] = "success"
        return {"data": data, "sources": [], "error": False, "debug": debug_info} # Sources handled via evidence array now
        
    except Exception as e:
        error_msg = str(e)
        debug_info["exception"] = error_msg
        debug_info["error_type"] = "Python Exception"
        if "429" in error_msg:
            return {"data": None, "sources": [], "error": True, "msg": "Quota exceeded.", "debug": debug_info}
        return {"data": None, "sources": [], "error": True, "msg": f"Error: {error_msg[:200]}", "debug": debug_info}

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}

st.title("🏛️ AHJ Research Assistant v9")
st.caption("Compact evidence-backed schema. Python validation. Dynamic disciplines.")

with st.sidebar:
    st.warning("⚠️ Pay-As-You-Go Active. Results cached for 1 hour.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=False)
    
    st.subheader("🐛 API Debug Log")
    st.json(st.session_state.debug_log)

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
    existing_permit = st.text_input("Existing Entitlements (Optional)", "e.g., Existing CUP, Variance #123")

# --- SECTION 2: SCOPE OF WORK ---
st.header("2. Scope of Work (SOW)")
sow_text = st.text_area(
    "Paste the complete Scope of Work below.", 
    height=150,
    value="Ground-level exterior HVAC unit replacement (like-for-like replacement with a different brand on an existing exterior pad; parcel operates under an existing Conditional Use Permit). should be in exact same place, ductwork wont be affected, no roof penetrations"
)

# --- SECTION 3: RESEARCH ---
st.header("3. Research Execution")

input_string = f"{PROMPT_VERSION}|{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    if mock_mode:
        st.session_state.report_data = {
            "bottom_line": "Mechanical permit is required. Electrical applicability depends on wiring changes. Existing CUP conditions must be verified.",
            "questions": [
                "Which agency has building jurisdiction for this parcel?",
                "Will the electrical disconnect, circuit, breaker, or conductors change?",
                "What are the replacement unit's capacity, weight, and efficiency ratings?"
            ],
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro (Unincorporated)", "ahj": "Washington County Dept. of Land Use (if unincorporated)", "evidence_ids": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence_ids": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "rule": "Jurisdiction for unincorporated areas."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "rule": "Commercial mechanical permit required for replacement."}
            ],
            "disciplines": [
                {
                    "type": "Mechanical", "permit": "VERIFIED_REQUIRED", "pathway": "CONDITIONAL",
                    "finding": "Commercial mechanical permit required.",
                    "evidence_ids": ["E2"],
                    "test": {"rule": "Permit required for regulated mechanical replacement.", "fact": "Commercial HVAC replacement.", "determination": "applies", "missing": "Unit specifications."},
                    "reopen": ""
                },
                {
                    "type": "Electrical", "permit": "CONDITIONAL", "pathway": "CONDITIONAL",
                    "finding": "Permit required only if electrical work is modified.",
                    "evidence_ids": ["E3"],
                    "test": {"rule": "Permits required for branch circuit/disconnect modification.", "fact": "SOW states 'different brand', electrical scope unstated.", "determination": "cannot_determine", "missing": "Unit electrical specs."},
                    "reopen": "Modifying electrical disconnect, wiring, or breaker amperage."
                }
            ]
        }
        st.session_state.debug_log = {"mock": True, "note": "No API call made"}
        st.info("🛡️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.error("GEMINI_KEY missing.")
        else:
            with st.spinner("Searching live databases and performing applicability tests..."):
                prompt = f"""
You are an expert AHJ research analyst. Perform targeted, authoritative research for this project and return ONLY valid JSON.

PROJECT:
State: {state}
Address: {address}
Date: {project_date}
Project Type: {ptype}
Building Class: {bclass} (USER-PROVIDED; do not independently assume it)
Existing Entitlements: {existing_permit}

SCOPE OF WORK:
{sow_text}

RESEARCH OBJECTIVE:
Determine jurisdiction, current applicable code editions, permit requirements, review pathways, conditional triggers, and unresolved facts.

SOURCE PRIORITY:
1. Official state agency
2. Official county/city/AHJ
3. Official permit portal
4. Official ordinance/code/interpretation

CRITICAL RULES:
1. JURISDICTION: Establish the actual building/AHJ jurisdiction from authoritative evidence. If it cannot be established, use CONDITIONAL.
2. CURRENT CODES: Use the current code edition applicable on the project date. NEVER invent adoption dates, effective dates, or mandatory dates. Every reported code must have an evidence_id.
3. PERMIT VS PATHWAY: Keep permit requirement separate from review pathway. If authoritative evidence establishes that a permit is required, permit must be VERIFIED_REQUIRED even when the pathway is conditional.
4. EVIDENCE: A VERIFIED conclusion requires authoritative retrieved evidence supporting the specific conclusion. A source homepage alone does not establish a detailed requirement.
5. APPLICABILITY: For each included discipline determine: what authoritative rule was found, what project fact applies, and whether applicability is established. If a material fact is missing, determination must be cannot_determine.
6. STATUS: Use ONLY: VERIFIED_REQUIRED, CONDITIONAL, INFERRED, UNKNOWN, NOT_APPLICABLE, NOT_CURRENTLY_TRIGGERED, USER_PROVIDED.
7. NOT_CURRENTLY_TRIGGERED: Use when the current scope does not trigger the discipline, but a specific new fact could reopen it. Do NOT use NOT_APPLICABLE when a missing fact could change the result.
8. DYNAMIC DISCIPLINES: Include disciplines that are currently triggered, conditionally triggered, materially unresolved, or reasonably important to reopening the analysis. Omit clearly irrelevant disciplines.
9. DO NOT SPECULATE: Do not invent permit portals, licensing requirements, code adoption dates, thresholds, or requirements unless supported by retrieved evidence.
10. BREVITY: Keep each string under 20 words. Do not repeat the same information across fields. Do not write prose paragraphs.
11. HIGH-VALUE QUESTIONS: Return no more than 5 questions. Only ask questions whose answers could materially change the research result.

RETURN ONLY THIS JSON:
{{
  "bottom_line": "...",
  "questions": ["...", "...", "..."],
  "jurisdiction": {{
    "status": "VERIFIED or CONDITIONAL",
    "county": "...",
    "city": "...",
    "ahj": "...",
    "evidence_ids": ["E1"]
  }},
  "codes": [
    {{
      "name": "...",
      "status": "CURRENT or CONDITIONAL",
      "evidence_ids": ["E2"]
    }}
  ],
  "evidence": [
    {{
      "id": "E1",
      "title": "...",
      "url": "...",
      "rule": "..."
    }}
  ],
  "disciplines": [
    {{
      "type": "Mechanical",
      "permit": "VERIFIED_REQUIRED",
      "pathway": "CONDITIONAL",
      "finding": "...",
      "evidence_ids": ["E2"],
      "test": {{
        "rule": "...",
        "fact": "...",
        "determination": "applies, does_not_apply, or cannot_determine",
        "missing": "..."
      }},
      "reopen": "..."
    }}
  ]
}}

IMPORTANT: Output JSON only. Do not use markdown. Do not add commentary. Do not output fields not specified above.
"""
                result = cached_gemini_call(prompt_hash, prompt)
                st.session_state.debug_log = result.get("debug", {})
                
                if result["error"]:
                    st.error(f"❌ {result['msg']}")
                else:
                    st.session_state.report_data = result["data"]
                    st.success("✅ Research dossier complete.")

# ============================================================
# RESULTS DISPLAY
# ============================================================
if st.session_state.report_data:
    data = st.session_state.report_data
    st.divider()
    
    st.header("4. Research Dossier")
    
    # Show validation warnings if any
    if "validation_errors" in st.session_state.debug_log:
        st.warning("⚠️ Schema Validation Warnings: " + " | ".join(st.session_state.debug_log["validation_errors"]))

    # 1. Bottom Line Summary
    st.info(f"**Bottom Line:** {data.get('bottom_line', 'N/A')}")

    # 2. Immediate Questions
    questions = data.get("questions") or []
    if questions:
        st.subheader("❓ Immediate Questions (Information Needed)")
        st.warning("Clarify these high-value facts to finalize permit pathways:")
        for i, q in enumerate(questions, 1):
            st.markdown(f"{i}. **{q}**")

    # 3. Jurisdiction
    st.subheader("📍 Jurisdiction Determination")
    jur = data.get("jurisdiction") or {}
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**County:** {jur.get('county', 'Unknown')}")
        st.write(f"**City:** {jur.get('city', 'Unknown')}")
        st.write(f"**Building AHJ:** {jur.get('ahj', 'Unknown')}")
    with col2:
        # Map evidence IDs to URLs for jurisdiction
        jur_ev_ids = jur.get("evidence_ids", [])
        ev_dict = {ev["id"]: ev for ev in data.get("evidence", [])}
        for eid in jur_ev_ids:
            if eid in ev_dict:
                st.write(f"**Source:** [{ev_dict[eid]['title']}]({ev_dict[eid]['url']})")

    # 4. Codes
    st.subheader("📚 Applicable Codes & Editions")
    ev_dict = {ev["id"]: ev for ev in data.get("evidence", [])}
    for code in (data.get("codes") or []):
        st.markdown(f"- **{code.get('name', 'Unknown')}** ({code.get('status', 'N/A')})")
        for eid in code.get("evidence_ids", []):
            if eid in ev_dict:
                st.markdown(f"  - *Source:* [{ev_dict[eid]['title']}]({ev_dict[eid]['url']})")

    # 5. Disciplines (Permit Matrix)
    st.subheader("📋 Permit & Review Matrix")
    
    status_map = {
        "VERIFIED_REQUIRED": "", "INFERRED": "🟡", "CONDITIONAL": "🟠",
        "UNKNOWN": "🔴", "NOT_APPLICABLE": "⚪", "NOT_CURRENTLY_TRIGGERED": "⚪", "USER_PROVIDED": ""
    }

    for item in data.get("disciplines", []):
        permit_emoji = status_map.get(item.get("permit", "UNKNOWN"), "")
        pathway_emoji = status_map.get(item.get("pathway", "UNKNOWN"), "")
        expander_title = f"{permit_emoji} {item.get('type', 'Unknown')} — Permit: {item.get('permit', 'UNKNOWN')} | Pathway: {pathway_emoji} {item.get('pathway', 'UNKNOWN')}"
        
        with st.expander(expander_title, expanded=False):
            st.write(f"**Finding:** {item.get('finding', 'N/A')}")
            
            # Map evidence IDs
            ev_ids = item.get("evidence_ids", [])
            if ev_ids:
                st.write("**Evidence:**")
                for eid in ev_ids:
                    if eid in ev_dict:
                        st.markdown(f"- [{ev_dict[eid]['title']}]({ev_dict[eid]['url']})")
            else:
                st.write("**Evidence:** None retrieved.")
            
            st.divider()
            st.subheader("Applicability Test")
            app_test = item.get("test") or {}
            st.write(f"**Rule:** {app_test.get('rule', 'N/A')}")
            st.write(f"**Fact:** {app_test.get('fact', 'N/A')}")
            st.write(f"**Determination:** {app_test.get('determination', 'N/A').replace('_', ' ').title()}")
            if app_test.get('missing'):
                st.info(f"**❓ Missing:** {app_test.get('missing')}")
            
            reopen = item.get('reopen', '')
            if reopen:
                st.success(f"**🔄 Reopen if:** {reopen}")

    # 6. Export
    st.header("5. Export")
    col1, col2 = st.columns(2)
    
    with col1:
        doc = Document()
        doc.styles["Normal"].font.name = "Aptos"
        doc.add_heading("AHJ Research Dossier", 0)
        doc.add_paragraph(f"Project: {address} ({state})\nDate: {project_date}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
        
        doc.add_heading("Bottom Line", level=1)
        doc.add_paragraph(data.get("bottom_line", ""))
        
        doc.add_heading("Immediate Questions", level=1)
        for i, q in enumerate(data.get("questions") or [], 1):
            doc.add_paragraph(f"{i}. {q}")
        
        doc.add_heading("Jurisdiction", level=1)
        jur = data.get("jurisdiction") or {}
        doc.add_paragraph(f"Status: {jur.get('status', 'N/A')}\nCounty: {jur.get('county', 'N/A')}\nCity: {jur.get('city', 'N/A')}\nAHJ: {jur.get('ahj', 'N/A')}")
        
        doc.add_heading("Applicable Codes", level=1)
        for code in (data.get("codes") or []): 
            doc.add_paragraph(f"{code.get('name')} ({code.get('status')})", style='List Bullet')
        
        doc.add_heading("Permit Matrix", level=1)
        for item in data.get("disciplines", []):
            doc.add_heading(f"{item.get('type')} - Permit: {item.get('permit')} | Pathway: {item.get('pathway')}", level=2)
            doc.add_paragraph(f"Finding: {item.get('finding')}")
            
            app_test = item.get("test") or {}
            doc.add_paragraph(f"Rule: {app_test.get('rule')}")
            doc.add_paragraph(f"Fact: {app_test.get('fact')}")
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
        json_data = json.dumps({
            "project": {"state": state, "address": address, "date": str(project_date)},
            "dossier": data
        }, indent=2)
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Dossier.json", mime="application/json", use_container_width=True)
