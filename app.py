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

st.set_page_config(page_title="AHJ Research Assistant v2", page_icon="🏛️", layout="wide")

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

# ============================================================
# CACHING & API CALL (WITH DEBUG LOG)
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0) # RPM throttle
    debug_info = {}
    
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        
        # REMOVED: response_mime_type="application/json" 
        # This was causing the model to choke when combined with Search Grounding.
        config = types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())],
            max_output_tokens=4000, 
        )

        response = client.models.generate_content(
            model="gemini-3.6-flash", 
            contents=prompt_text,
            config=config,
        )
        
        # 1. Capture Debug Info
        if response.candidates:
            debug_info["finish_reason"] = str(response.candidates[0].finish_reason)
            debug_info["safety_ratings"] = [str(r) for r in response.candidates[0].safety_ratings]
            
            finish_reason = response.candidates[0].finish_reason
            if finish_reason and str(finish_reason) != "FinishReason.STOP":
                if "SAFETY" in str(finish_reason):
                    return {"data": None, "sources": [], "error": True, "msg": "Blocked by Safety Filter.", "debug": debug_info}
                return {"data": None, "sources": [], "error": True, "msg": f"API stopped early: {finish_reason}", "debug": debug_info}

        # 2. Safely extract text
        text = getattr(response, "text", None)
        if not text:
            try:
                text = response.candidates[0].content.parts[0].text
            except Exception:
                pass
                
        if not text:
            return {"data": None, "sources": [], "error": True, "msg": "Empty response.", "debug": debug_info}

        # 3. Bulletproof JSON Parsing (Regex)
        data = None
        try:
            # Find the largest JSON object in the text
            match = re.search(r'\{.*\}', text, re.DOTALL)
            if match:
                data = json.loads(match.group(0))
            else:
                raise ValueError("No JSON object found")
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:200]
            return {"data": None, "sources": [], "error": True, "msg": "Failed to parse JSON.", "debug": debug_info}

        # 4. Extract sources
        sources = []
        try:
            if response.candidates:
                metadata = getattr(response.candidates[0], "grounding_metadata", None)
                if metadata and getattr(metadata, "grounding_chunks", None):
                    for chunk in metadata.grounding_chunks:
                        if getattr(chunk, "web", None):
                            sources.append({"title": chunk.web.title, "url": chunk.web.uri})
        except Exception:
            pass
            
        return {"data": data, "sources": sources, "error": False, "debug": debug_info}
        
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg:
            return {"data": None, "sources": [], "error": True, "msg": "Quota exceeded.", "debug": {}}
        return {"data": None, "sources": [], "error": True, "msg": f"Error: {error_msg[:200]}", "debug": {}}

# ============================================================
# UI & STATE
# ============================================================
if "report_data" not in st.session_state: st.session_state.report_data = None
if "sources" not in st.session_state: st.session_state.sources = []
if "debug_log" not in st.session_state: st.session_state.debug_log = None

st.title("🏛️ AHJ Research Assistant v2")
st.caption("Evidence-first architecture. Structured research dossier. Fail-safe confidence.")

with st.sidebar:
    st.warning("️ Pay-As-You-Go Active. Results cached for 1 hour.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=False)
    
    # DEBUG LOG EXPANDER
    if st.session_state.debug_log:
        with st.expander("🐛 API Debug Log", expanded=True):
            st.json(st.session_state.debug_log)
            
    if st.session_state.sources:
        st.success(f"✅ {len(st.session_state.sources)} live sources found")

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
    "Paste the complete Scope of Work below. Do not summarize.", 
    height=250,
    value="Ground-level exterior HVAC unit replacement (like-for-like replacement with a different brand on an existing exterior pad; parcel operates under an existing Conditional Use Permit). should be in exact same place, ductwork wont be affected, no roof penetrations"
)

# --- SECTION 3: RESEARCH ---
st.header("3. Research Execution")

input_string = f"{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    if mock_mode:
        st.session_state.report_data = {
            "user_scope_verbatim": sow_text,
            "research_interpretation": "Commercial HVAC replacement on existing pad.",
            "detected_categories": ["HVAC / Mechanical", "Electrical / Power"],
            "permit_matrix": [
                {"permit_type": "Mechanical", "result": "Likely Required", "evidence_status": "VERIFIED", "why": "State code requires permits for commercial HVAC replacement.", "applicability": "Applies to all commercial mechanical work.", "evidence_sources": ["Oregon BCD Mechanical Code"], "unknowns": "None"},
                {"permit_type": "Electrical", "result": "Conditional", "evidence_status": "CONDITIONAL", "why": "Depends on if wiring is modified.", "applicability": "Only if disconnect is moved.", "evidence_sources": [], "unknowns": "Will wiring be modified?"}
            ],
            "applicable_codes": ["2025 Oregon Mechanical Specialty Code"],
            "hidden_triggers": ["Verify refrigerant type for A2L compliance."],
            "action_plan": ["1. Submit mechanical permit application.", "2. Schedule inspection."]
        }
        st.session_state.sources = [{"title": "Mock Source", "url": "https://example.com"}]
        st.session_state.debug_log = {"mock": True}
        st.info("🛡️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.error("GEMINI_KEY missing.")
        else:
            with st.spinner("Searching live databases and building evidence dossier..."):
                prompt = f"""
You are an expert AHJ (Authority Having Jurisdiction) research assistant. You MUST output a STRICT JSON object. Do not include conversational text outside the JSON block.

PROJECT METADATA:
- State: {state}
- Address: {address}
- Date: {project_date}
- Project Type: {ptype}
- Building Class: {bclass}
- Existing Entitlements: {existing_permit}

USER-STATED SCOPE OF WORK:
{sow_text}

CRITICAL RESEARCH RULES:
1. Do not treat a search-result snippet as evidence. Verify the complete provision.
2. Do not infer a legal exemption merely because you found a threshold. Verify applicability.
3. Never create a citation to a source you did not retrieve.
4. If you cannot establish a conclusion from authoritative evidence, return "UNKNOWN" rather than guessing.
5. Ignore cosmetic fluff (paint, carpet). Focus on structural, mechanical, electrical, plumbing, fire, and zoning triggers.

OUTPUT JSON SCHEMA (Strictly follow this structure):
{{
  "user_scope_verbatim": "The exact text provided by the user.",
  "research_interpretation": "A concise technical summary of the actual work being performed.",
  "detected_categories": ["List", "of", "applicable", "trade", "categories"],
  "permit_matrix": [
    {{
      "permit_type": "Mechanical",
      "result": "Likely Required",
      "evidence_status": "VERIFIED", 
      "why": "Brief explanation of the rule.",
      "applicability": "How it applies to this specific scope.",
      "evidence_sources": ["Source Name 1", "Source Name 2"],
      "unknowns": "What is still unknown or needs verification."
    }}
  ],
  "applicable_codes": ["List of specific code editions and effective dates."],
  "hidden_triggers": ["List of missing details or clarifying questions."],
  "action_plan": ["Chronological step-by-step list."]
}}

EVIDENCE STATUS MUST BE ONE OF:
- "VERIFIED": Authoritative source retrieved and explicitly supports the conclusion.
- "INFERRED": Evidence exists, but conclusion requires professional interpretation.
- "CONDITIONAL": True only if an unresolved fact is true.
- "UNKNOWN": Evidence was not found.
- "SCOPE_BASED": Conclusion derived purely from user scope, no external evidence needed.
"""
                result = cached_gemini_call(prompt_hash, prompt)
                st.session_state.debug_log = result.get("debug", {})
                
                if result["error"]:
                    st.error(f"❌ {result['msg']} (Check Debug Log in sidebar)")
                else:
                    st.session_state.report_data = result["data"]
                    st.session_state.sources = result["sources"]
                    st.success("✅ Research dossier complete.")

# ============================================================
# RESULTS DISPLAY (DETERMINISTIC UI)
# ============================================================
if st.session_state.report_data:
    data = st.session_state.report_data
    st.divider()
    
    st.header("4. Research Dossier")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📝 User-Stated Scope")
        st.info(data.get("user_scope_verbatim", "N/A"))
    with col2:
        st.subheader("🔍 Research Interpretation")
        st.success(data.get("research_interpretation", "N/A"))
        
    st.caption(f"**AI Detected Categories:** {', '.join(data.get('detected_categories', []))}")

    st.subheader("Permit & Review Matrix")
    
    status_map = {
        "VERIFIED": "🟢",
        "INFERRED": "🟡",
        "CONDITIONAL": "🟠",
        "UNKNOWN": "🔴",
        "SCOPE_BASED": "⚪"
    }

    for item in data.get("permit_matrix", []):
        status_emoji = status_map.get(item.get("evidence_status", "UNKNOWN"), "")
        expander_title = f"{item.get('permit_type', 'Unknown')} — {item.get('result', 'Unknown')} ({status_emoji} {item.get('evidence_status', 'UNKNOWN')})"
        
        with st.expander(expander_title, expanded=False):
            st.write(f"**Why:** {item.get('why', 'N/A')}")
            st.write(f"**Applicability:** {item.get('applicability', 'N/A')}")
            
            sources = item.get('evidence_sources', [])
            if sources:
                st.write(f"**Evidence:** {', '.join(sources)}")
            else:
                st.warning("**Evidence:** None retrieved.")
                
            unknowns = item.get('unknowns', '')
            if unknowns and unknowns.lower() != 'none':
                st.error(f"**️ Unknowns / To Verify:** {unknowns}")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Applicable Codes & Editions")
        for code in data.get("applicable_codes", []):
            st.markdown(f"- {code}")
    with col2:
        st.subheader("Hidden Triggers & Questions")
        for trigger in data.get("hidden_triggers", []):
            st.markdown(f"- {trigger}")

    st.subheader("Volunteer Action Plan")
    for i, step in enumerate(data.get("action_plan", []), 1):
        st.markdown(f"{i}. {step}")

    if st.session_state.sources:
        with st.expander(" Live Sources Retrieved", expanded=False):
            for i, s in enumerate(st.session_state.sources, 1):
                st.markdown(f"**{i}.** [{s['title']}]({s['url']})\n   `{s['url']}`")

    st.header("5. Export")
    col1, col2 = st.columns(2)
    
    with col1:
        doc = Document()
        doc.styles["Normal"].font.name = "Aptos"
        doc.add_heading("AHJ Research Dossier", 0)
        doc.add_paragraph(f"Project: {address} ({state})\nDate: {project_date}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
        
        doc.add_heading("User-Stated Scope", level=1)
        doc.add_paragraph(data.get("user_scope_verbatim", ""))
        
        doc.add_heading("Research Interpretation", level=1)
        doc.add_paragraph(data.get("research_interpretation", ""))
        
        doc.add_heading("Permit Matrix", level=1)
        for item in data.get("permit_matrix", []):
            doc.add_heading(f"{item.get('permit_type')} - {item.get('result')} [{item.get('evidence_status')}]", level=2)
            doc.add_paragraph(f"Why: {item.get('why')}")
            doc.add_paragraph(f"Applicability: {item.get('applicability')}")
            doc.add_paragraph(f"Evidence: {', '.join(item.get('evidence_sources', []))}")
            if item.get('unknowns'): doc.add_paragraph(f"Unknowns: {item.get('unknowns')}")
            
        doc.add_heading("Applicable Codes", level=1)
        for code in data.get("applicable_codes", []): doc.add_paragraph(code, style='List Bullet')
        
        doc.add_heading("Hidden Triggers", level=1)
        for trigger in data.get("hidden_triggers", []): doc.add_paragraph(trigger, style='List Bullet')
        
        doc.add_heading("Action Plan", level=1)
        for i, step in enumerate(data.get("action_plan", []), 1): doc.add_paragraph(f"{i}. {step}")
            
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📄 Download Word Report", data=buf.getvalue(), file_name="AHJ_Dossier.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    with col2:
        json_data = json.dumps({
            "project": {"state": state, "address": address, "date": str(project_date)},
            "dossier": data,
            "sources": st.session_state.sources
        }, indent=2)
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Dossier.json", mime="application/json", use_container_width=True)
