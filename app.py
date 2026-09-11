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

st.set_page_config(page_title="AHJ Research Assistant", page_icon="🏛️", layout="wide")

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

PROJECT_TYPES = [
    "Replacement / Repair", "Remodel / Tenant Improvement", 
    "Addition", "New Construction", "Site / Civil Work", "Other"
]

BUILDING_CLASSES = [
    "Commercial", "Assembly", "Institutional", "Industrial", 
    "Agricultural", "Residential (1-2 Family)", "Residential (Multi-family)", 
    "Mixed-use", "Unknown"
]

SCOPE_CATEGORIES = [
    "HVAC / Mechanical", "Roofing / Envelope", "Electrical / Power", 
    "Plumbing / Fire Protection", "Site Work / Civil / Parking", 
    "Structural / Foundation", "Interior / Architectural", 
    "Energy / Sustainability", "Fire / Life Safety", "Zoning / Land Use"
]

GEMINI_KEY = os.getenv("GEMINI_KEY") or st.secrets.get("GEMINI_KEY", "")

# ============================================================
# CACHING & API CALL
# ============================================================
@st.cache_data(ttl=3600)
def cached_gemini_call(prompt_hash, prompt_text):
    time.sleep(1.0) # Throttle to respect RPM limits
    
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        response = client.models.generate_content(
            model="gemini-3.6-flash", 
            contents=prompt_text,
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
                max_output_tokens=4000, # Generous limit for massive scopes
            )
        )
        
        sources = []
        try:
            for chunk in response.candidates[0].grounding_metadata.grounding_chunks:
                if getattr(chunk, "web", None):
                    sources.append({"title": chunk.web.title, "url": chunk.web.uri})
        except Exception:
            pass
            
        return {"text": response.text.strip(), "sources": sources, "error": False}
        
    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg:
            return {"text": "ERROR: Quota exceeded. Please wait a moment.", "sources": [], "error": True}
        return {"text": f"ERROR: {error_msg[:200]}", "sources": [], "error": True}

# ============================================================
# UI & STATE
# ============================================================
if "report" not in st.session_state: st.session_state.report = ""
if "sources" not in st.session_state: st.session_state.sources = []

st.title("🏛️ AHJ Research Assistant")
st.caption("Dynamic Scope Analysis. Live Code & Permit Research.")

with st.sidebar:
    st.warning("️ Pay-As-You-Go Active. Results cached for 1 hour.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=False)
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

# --- SECTION 2: SMART SCOPE BUILDER ---
st.header("2. Scope of Work (SOW)")
st.info("Select all categories that apply, then paste the detailed Scope of Work below. The AI will filter out cosmetic fluff and focus on permit triggers.")

selected_categories = st.multiselect(
    "Applicable Scope Categories", 
    SCOPE_CATEGORIES, 
    default=["HVAC / Mechanical"]
)

# Made taller for massive copy-pastes
sow_text = st.text_area(
    "Detailed Scope of Work (Paste entire document)", 
    height=300,
    value="Ground-level exterior HVAC unit replacement (like-for-like replacement with a different brand on an existing exterior pad). Unit will be in exact same location. Ductwork will not be affected. No roof penetrations. Parcel operates under an existing Conditional Use Permit."
)

# --- SECTION 3: RESEARCH ---
st.header("3. Research Execution")

input_string = f"{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{','.join(selected_categories)}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Run AHJ Research", type="primary", use_container_width=True):
    if mock_mode:
        st.session_state.report = "# MOCK REPORT\n\nThis is a mock report to test UI without burning API credits."
        st.session_state.sources = [{"title": "Mock Source", "url": "https://example.com"}]
        st.info("️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.error("GEMINI_KEY missing.")
        else:
            with st.spinner("Analyzing scope and searching live government databases..."):
                prompt = f"""
You are an expert AHJ (Authority Having Jurisdiction) and Building Code research assistant.

PROJECT METADATA:
- State: {state}
- Address: {address}
- Date: {project_date}
- Project Type: {ptype}
- Building Class: {bclass}
- Existing Entitlements: {existing_permit}

SCOPE CATEGORIES: {', '.join(selected_categories) if selected_categories else 'None specified'}

DETAILED SCOPE OF WORK:
{sow_text}

CRITICAL INSTRUCTIONS FOR ANALYSIS:
1. SCOPE SIZE ADAPTATION: If the SOW is very large, ignore cosmetic updates (paint, carpet, 'retain', 'branch standard'). Focus STRICTLY on structural, mechanical, electrical, plumbing, fire, and zoning triggers. If the SOW is small, be comprehensive.
2. LIVE SEARCH: Use your Google Search tool to find current, official government sources for this specific jurisdiction and date.

OUTPUT FORMAT (Strictly follow this structure):

# 1. EXECUTIVE SUMMARY & PERMIT MATRIX
(Brief 2-3 sentence summary of the project's path).
| Permit / Review Type | Required? | Triggering Factor | Confidence |
|---|---|---|---|
(List all relevant permits: Building, Mechanical, Electrical, Plumbing, Fire, Planning/Zoning, Public Works. Mark as Yes, No, or Conditional).

# 2. APPLICABLE CODES & EDITIONS
CRITICAL RULE: You MUST list the specific adopted code edition for EVERY permit marked "Yes" or "Conditional" in Section 1. Do not skip any.
- Building Code: [Edition, Effective Date, Source URL]
- Mechanical Code: [Edition, Effective Date, Source URL]
- Electrical Code: [Edition, Effective Date, Source URL]
- Plumbing Code: [Edition, Effective Date, Source URL]
- Energy Code: [Edition, Effective Date, Source URL]
- Existing Building Code: [Edition, Effective Date, Source URL] (Only if applicable)

# 3. SCOPE-SPECIFIC REQUIREMENTS
(Break down the requirements based on the selected categories and SOW. Be specific about thresholds, e.g., "Mechanical permit required for any replacement > X BTU").

# 4. HIDDEN TRIGGERS & CLARIFYING QUESTIONS
(This is critical. Look at the SOW and identify what is MISSING. If they mention roofing but not structural deck condition, ask about it. List 3-5 precise questions the volunteer MUST answer or ask the AHJ).

# 5. VOLUNTEER ACTION PLAN
(Chronological step-by-step list to get this project approved).

# 6. SOURCES
(List titles and URLs of official .gov sources found).
"""
                result = cached_gemini_call(prompt_hash, prompt)
                
                if result["error"]:
                    st.error(result["text"])
                else:
                    st.session_state.report = result["text"]
                    st.session_state.sources = result["sources"]
                    st.success("✅ Research complete.")

# ============================================================
# RESULTS DISPLAY
# ============================================================
if st.session_state.report:
    st.divider()
    st.header("4. Research Report")
    st.markdown(st.session_state.report)
    
    if st.session_state.sources:
        with st.expander("🔗 Live Sources Retrieved", expanded=False):
            for i, s in enumerate(st.session_state.sources, 1):
                st.markdown(f"**{i}.** [{s['title']}]({s['url']})\n   `{s['url']}`")

    st.header("5. Export")
    col1, col2 = st.columns(2)
    
    with col1:
        doc = Document()
        doc.styles["Normal"].font.name = "Aptos"
        doc.add_heading("AHJ Research Report", 0)
        doc.add_paragraph(f"Project: {address} ({state})\nDate: {project_date}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
        doc.add_heading("Report", level=1)
        
        for line in st.session_state.report.splitlines():
            line = line.strip()
            if not line: continue
            if line.startswith("# "): doc.add_heading(line[2:], level=1)
            elif line.startswith("## "): doc.add_heading(line[3:], level=2)
            elif line.startswith("- "): doc.add_paragraph(line[2:], style="List Bullet")
            elif line.startswith("|"): doc.add_paragraph(line)
            else: doc.add_paragraph(line)
            
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📄 Download Word Report", data=buf.getvalue(), file_name="AHJ_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    with col2:
        json_data = json.dumps({
            "project": {"state": state, "address": address, "date": str(project_date)},
            "report": st.session_state.report,
            "sources": st.session_state.sources
        }, indent=2)
        st.download_button("💾 Save JSON Session", data=json_data, file_name="AHJ_Session.json", mime="application/json", use_container_width=True)
