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

st.set_page_config(page_title="AHJ Research Assistant v7", page_icon="🏛️", layout="wide")

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
            
            # EXPLICIT MAX_TOKENS HANDLING
            if finish_reason == "FinishReason.MAX_TOKENS":
                debug_info["error_type"] = "MAX_TOKENS"
                return {"data": None, "sources": [], "error": True, "msg": "Model hit output limit. Try a shorter SOW or simplify the request.", "debug": debug_info}
                
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
            match = re.search(r'\{.*\}', text, re.DOTALL)
            if match:
                data = json.loads(match.group(0))
            else:
                raise ValueError("No JSON object found")
        except Exception as e:
            debug_info["json_error"] = str(e)
            debug_info["raw_text_snippet"] = text[:500]
            debug_info["error_type"] = "JSON Parse Failed"
            return {"data": None, "sources": [], "error": True, "msg": "Failed to parse JSON.", "debug": debug_info}

        sources = []
        try:
            metadata = getattr(response.candidates[0], "grounding_metadata", None)
            if metadata and getattr(metadata, "grounding_chunks", None):
                for chunk in metadata.grounding_chunks:
                    if getattr(chunk, "web", None):
                        sources.append({"title": chunk.web.title, "url": chunk.web.uri})
        except Exception:
            pass
            
        debug_info["status"] = "success"
        return {"data": data, "sources": sources, "error": False, "debug": debug_info}
        
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
if "sources" not in st.session_state: st.session_state.sources = []
if "debug_log" not in st.session_state: st.session_state.debug_log = {"status": "Waiting for first run..."}

st.title("🏛️ AHJ Research Assistant v7")
st.caption("Compact schema. Dynamic disciplines. Single-pass token efficiency.")

with st.sidebar:
    st.warning("⚠️ Pay-As-You-Go Active. Results cached for 1 hour.")
    mock_mode = st.toggle("🛡️ Mock Mode", value=False)
    
    st.subheader("🐛 API Debug Log")
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
    "Paste the complete Scope of Work below.", 
    height=150,
    value="Ground-level exterior HVAC unit replacement (like-for-like replacement with a different brand on an existing exterior pad; parcel operates under an existing Conditional Use Permit). should be in exact same place, ductwork wont be affected, no roof penetrations"
)

# --- SECTION 3: RESEARCH ---
st.header("3. Research Execution")

input_string = f"{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    if mock_mode:
        st.session_state.report_data = {
            "bottom_line": "Mechanical: VERIFIED REQUIRED (pathway conditional). Electrical: CONDITIONAL. Planning: CONDITIONAL (CUP).",
            "questions": ["What is the new unit's weight and CFM?", "Will the electrical disconnect or branch circuit be modified?", "Does the CUP have specific screening/noise conditions?"],
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro (Unincorporated)", "ahj": "Washington County Dept. of Land Use", "portal": "https://www.washingtoncounty.org/1134/Building-Services"},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code (OMSC)", "edition": "2025", "mandatory": "April 1, 2026"}],
            "disciplines": [
                {
                    "type": "Mechanical", "status": "VERIFIED", "pathway": "CONDITIONAL",
                    "finding": "Commercial mechanical permit required for replacement.",
                    "evidence": "Washington County Commercial Building Page",
                    "gap": "Does not establish minor-installation exemption eligibility.",
                    "needs": "Proposed unit weight, CFM, and cooling capacity.",
                    "test": {"rule": "Permit required for regulated system replacement.", "fact": "Replacing commercial HVAC.", "compare": "Rule applies, but pathway depends on specs.", "determination": "cannot_determine", "missing": "Unit specs."},
                    "reopen": "N/A"
                },
                {
                    "type": "Electrical", "status": "CONDITIONAL", "pathway": "CONDITIONAL",
                    "finding": "Permit depends on whether electrical work is modified.",
                    "evidence": "OESC 105.1",
                    "gap": "Does not establish if existing circuit is compatible.",
                    "needs": "Existing and proposed unit MCA, MOP, voltage.",
                    "test": {"rule": "Permits required for branch circuit/disconnect modification.", "fact": "SOW states 'different brand', electrical scope unstated.", "compare": "If compatible and unchanged, no permit. If changed, permit required.", "determination": "cannot_determine", "missing": "Unit electrical specs."},
                    "reopen": "N/A"
                }
            ]
        }
        st.session_state.sources = [{"title": "Mock Source", "url": "https://example.com"}]
        st.session_state.debug_log = {"mock": True, "note": "No API call made"}
        st.info("🛡️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.error("GEMINI_KEY missing.")
        else:
            with st.spinner("Analyzing scope and generating targeted research..."):
                prompt = f"""
You are an expert AHJ research assistant. You MUST output a STRICT, COMPACT JSON object.

PROJECT METADATA:
- State: {state}
- Address: {address}
- Date: {project_date}
- Project Type: {ptype}
- Building Class: {bclass} (Treat as USER-PROVIDED metadata)
- Existing Entitlements: {existing_permit}

USER-STATED SCOPE OF WORK:
{sow_text}

CRITICAL RESEARCH & APPLICABILITY RULES:
1. DYNAMIC DISCIPLINES: ONLY include disciplines in the `disciplines` array that have a current trigger, a plausible conditional trigger, or a material unresolved question. Omit completely irrelevant disciplines (e.g., do not include Fire/Life Safety for a simple interior HVAC swap unless a specific trigger exists).
2. EXTREME BREVITY: Every string value MUST be under 15 words. Use telegraphic style. Do not repeat information across fields.
3. THE "NOT_APPLICABLE" BAN: You CANNOT output a discipline as "NOT_APPLICABLE" if a missing fact could change the outcome. Use "NOT_CURRENTLY_TRIGGERED" and provide a `reopen` condition.
4. JURISDICTION CONSISTENCY: If Jurisdiction status is "CONDITIONAL", downstream permits MUST reflect this (e.g., status: "CONDITIONAL").
5. APPLICABILITY TEST: For EVERY included discipline, fill out the `test` object. Use "cannot_determine" if a key fact is missing.

OUTPUT JSON SCHEMA (Strictly follow this structure. Keep values extremely short):
{{
  "bottom_line": "2-3 concise sentences stating the status of top disciplines.",
  "questions": ["Top 1 high-value question", "Top 2 high-value question", "Top 3 high-value question"],
  "jurisdiction": {{
    "status": "VERIFIED or CONDITIONAL",
    "county": "...",
    "city": "...",
    "ahj": "...",
    "portal": "..."
  }},
  "codes": [
    {{"name": "...", "edition": "...", "mandatory": "..."}}
  ],
  "disciplines": [
    {{
      "type": "Mechanical",
      "status": "CONDITIONAL", 
      "pathway": "CONDITIONAL",
      "finding": "Brief summary of the finding (max 15 words).",
      "evidence": "Source name or code section (max 10 words).",
      "gap": "Crucial gap remaining (max 15 words).",
      "needs": "Specific fact needed (max 15 words).",
      "test": {{
        "rule": "Specific rule found.",
        "fact": "Specific fact from SOW/metadata.",
        "compare": "How they compare.",
        "determination": "applies, does_not_apply, or cannot_determine",
        "missing": "What is missing."
      }},
      "reopen": "N/A, OR specific condition that would trigger this."
    }}
  ]
}}

ALLOWED STATUS VALUES: "VERIFIED", "CONDITIONAL", "INFERRED", "UNKNOWN", "NOT_APPLICABLE", "NOT_CURRENTLY_TRIGGERED", "USER_PROVIDED"
"""
                result = cached_gemini_call(prompt_hash, prompt)
                st.session_state.debug_log = result.get("debug", {})
                
                if result["error"]:
                    st.error(f"❌ {result['msg']}")
                else:
                    st.session_state.report_data = result["data"]
                    st.session_state.sources = result["sources"]
                    st.success("✅ Research dossier complete.")

# ============================================================
# RESULTS DISPLAY
# ============================================================
if st.session_state.report_data:
    data = st.session_state.report_data
    st.divider()
    
    st.header("4. Research Dossier")
    
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
    st.subheader(" Jurisdiction Determination")
    jur = data.get("jurisdiction") or {}
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**County:** {jur.get('county', 'Unknown')}")
        st.write(f"**City:** {jur.get('city', 'Unknown')}")
        st.write(f"**Building AHJ:** {jur.get('ahj', 'Unknown')}")
    with col2:
        portal = jur.get('portal', '')
        if portal: st.write(f"**Portal:** [Link]({portal})")

    # 4. Codes
    st.subheader(" Applicable Codes & Editions")
    for code in (data.get("codes") or []):
        st.markdown(f"- **{code.get('name', 'Unknown')}** (Edition: {code.get('edition', 'N/A')}, Mandatory: {code.get('mandatory', 'N/A')})")

    # 5. Disciplines (Permit Matrix)
    st.subheader("📋 Permit & Review Matrix")
    
    status_map = {
        "VERIFIED": "", "INFERRED": "🟡", "CONDITIONAL": "🟠",
        "UNKNOWN": "🔴", "NOT_APPLICABLE": "⚪", "NOT_CURRENTLY_TRIGGERED": "⚪", "USER_PROVIDED": ""
    }

    for item in data.get("disciplines", []):
        status_emoji = status_map.get(item.get("status", "UNKNOWN"), "")
        pathway_emoji = status_map.get(item.get("pathway", "UNKNOWN"), "")
        expander_title = f"{status_emoji} {item.get('type', 'Unknown')} — Permit: {item.get('status', 'UNKNOWN')} | Pathway: {pathway_emoji} {item.get('pathway', 'UNKNOWN')}"
        
        with st.expander(expander_title, expanded=False):
            st.write(f"**Finding:** {item.get('finding', 'N/A')}")
            st.write(f"**Evidence:** {item.get('evidence', 'None retrieved.')}")
            
            st.divider()
            st.warning(f"**⚠️ Gap:** {item.get('gap', 'Nothing.')}")
            st.info(f"**❓ Still Needs:** {item.get('needs', 'Nothing.')}")
            
            reopen = item.get('reopen', 'N/A')
            if reopen and reopen.lower() != 'n/a':
                st.success(f"**🔄 Reopen if:** {reopen}")
            
            st.subheader("Applicability Test Logic")
            app_test = item.get("test") or {}
            st.write(f"**Rule:** {app_test.get('rule', 'N/A')}")
            st.write(f"**Fact:** {app_test.get('fact', 'N/A')}")
            st.write(f"**Compare:** {app_test.get('compare', 'N/A')}")
            st.write(f"**Determination:** {app_test.get('determination', 'N/A').replace('_', ' ').title()}")
            st.write(f"**Missing:** {app_test.get('missing', 'N/A')}")

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
            doc.add_paragraph(f"{code.get('name')} ({code.get('edition')}) - Mandatory: {code.get('mandatory')}", style='List Bullet')
        
        doc.add_heading("Permit Matrix", level=1)
        for item in data.get("disciplines", []):
            doc.add_heading(f"{item.get('type')} - Permit: {item.get('status')} | Pathway: {item.get('pathway')}", level=2)
            doc.add_paragraph(f"Finding: {item.get('finding')}")
            doc.add_paragraph(f"Evidence: {item.get('evidence')}")
            doc.add_paragraph(f"Gap: {item.get('gap')}")
            doc.add_paragraph(f"Still needs: {item.get('needs')}")
            
            reopen = item.get('reopen', 'N/A')
            if reopen and reopen.lower() != 'n/a':
                doc.add_paragraph(f"Reopen if: {reopen}")
            
            app_test = item.get("test") or {}
            doc.add_paragraph(f"Rule: {app_test.get('rule')}")
            doc.add_paragraph(f"Fact: {app_test.get('fact')}")
            doc.add_paragraph(f"Determination: {app_test.get('determination', '').replace('_', ' ').title()}")
            doc.add_paragraph(f"Missing: {app_test.get('missing')}")
            
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
