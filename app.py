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

st.set_page_config(page_title="AHJ Research Assistant v4", page_icon="🏛️", layout="wide")

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
        
        # 1. Capture Debug Info SAFELY
        if response.candidates:
            candidate = response.candidates[0]
            debug_info["finish_reason"] = str(candidate.finish_reason)
            ratings = getattr(candidate, "safety_ratings", None) or []
            debug_info["safety_ratings"] = [str(r) for r in ratings]
            
            finish_reason = candidate.finish_reason
            if finish_reason and str(finish_reason) != "FinishReason.STOP":
                debug_info["error_type"] = "Early Stop"
                return {"data": None, "sources": [], "error": True, "msg": f"API stopped early: {finish_reason}", "debug": debug_info}
        else:
            debug_info["candidates"] = None
            debug_info["prompt_feedback"] = str(getattr(response, "prompt_feedback", None))
            debug_info["error_type"] = "No Candidates"
            return {"data": None, "sources": [], "error": True, "msg": "No candidates returned.", "debug": debug_info}

        # 2. Safely extract text
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

        # 3. Bulletproof JSON Parsing (Regex)
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

        # 4. Extract sources
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

st.title("🏛️ AHJ Research Assistant v4")
st.caption("Evidence-first architecture. Applicability testing. Permit vs. Pathway distinction.")

with st.sidebar:
    st.warning("️ Pay-As-You-Go Active. Results cached for 1 hour.")
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
            "bottom_line_summary": "A commercial mechanical permit is required for this HVAC replacement. The remaining research question is which mechanical permitting/review pathway applies, not whether a mechanical permit is required.",
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Unknown - requires parcel verification", "building_ahj": "Not yet established", "planning_ahj": "Not yet established", "permit_portal_url": ""},
            "applicable_codes": [{"code_name": "Oregon Mechanical Specialty Code (OMSC)", "edition": "2025", "mandatory_date": "April 1, 2026", "source_url": "https://www.oregon.gov/bcd"}],
            "permit_matrix": [
                {
                    "permit_type": "Mechanical",
                    "permit_status": "VERIFIED",
                    "review_pathway_status": "CONDITIONAL",
                    "evidence_quality": "Official AHJ checklist/application",
                    "summary": "Commercial HVAC replacement requires mechanical permitting.",
                    "why": "Hillsboro/Washington County requires permits for all commercial mechanical work. OMSC 105.1 requires permits to replace regulated systems.",
                    "evidence": "Washington County Commercial Building Page",
                    "applicability_test": {
                        "source_rule": "Commercial equipment exceeding specified weight/capacity thresholds may fall outside the minor mechanical installation category.",
                        "project_fact": "Proposed equipment weight and CFM have not been provided.",
                        "comparison": "Cannot compare proposed equipment to applicable minor-installation thresholds.",
                        "determination": "cannot_determine",
                        "missing_fact": "Proposed operating weight and CFM"
                    },
                    "what_i_still_need_from_you": "Proposed equipment specifications to determine the applicable permit/review pathway."
                },
                {
                    "permit_type": "Energy",
                    "permit_status": "CONDITIONAL",
                    "review_pathway_status": "CONDITIONAL",
                    "evidence_quality": "Search result only",
                    "summary": "Current energy requirements may apply to replacement mechanical equipment.",
                    "why": "2025 OEESC is mandatory, but specific replacement-equipment provisions must be verified.",
                    "evidence": "None retrieved.",
                    "applicability_test": {
                        "source_rule": "2025 OEESC contains provisions for replacement equipment.",
                        "project_fact": "Existing and proposed equipment efficiency ratings are unknown.",
                        "comparison": "Cannot determine if the replacement qualifies under an alteration/replacement provision.",
                        "determination": "cannot_determine",
                        "missing_fact": "Existing and proposed equipment type, capacity, and efficiency ratings."
                    },
                    "what_i_still_need_from_you": "Existing and proposed equipment efficiency ratings."
                }
            ],
            "hidden_triggers": ["Check if the existing CUP contains conditions governing exterior mechanical equipment."],
            "action_plan": ["1. Confirm exact jurisdiction (City vs. County).", "2. Gather proposed equipment specifications (Weight, CFM, MCA, MOCP).", "3. Determine whether seismic anchorage documentation is required by the AHJ."]
        }
        st.session_state.sources = [{"title": "Mock Source", "url": "https://example.com"}]
        st.session_state.debug_log = {"mock": True, "note": "No API call made"}
        st.info("️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.error("GEMINI_KEY missing.")
        else:
            with st.spinner("Searching live databases and performing applicability tests..."):
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

CRITICAL RESEARCH & APPLICABILITY RULES:
1. JURISDICTION HARD GATE: If you cannot definitively prove the exact City/County AHJ from the address, set jurisdiction status to "CONDITIONAL" and state "AHJ not yet confirmed". Do not guess between City and County.
2. PERMIT vs. PATHWAY DISTINCTION: "Is a permit required?" and "What is the review pathway?" are TWO DIFFERENT CONCLUSIONS. If a rule states commercial HVAC requires a permit, set `permit_status` to VERIFIED. If the project lacks weight/CFM to determine if it qualifies for a minor label, set `review_pathway_status` to CONDITIONAL. DO NOT downgrade a VERIFIED permit requirement just because the pathway is unknown.
3. APPLICABILITY TEST: For every finding, you MUST include an `applicability_test` object showing: Source Rule, Project Fact, Comparison, Determination (applies/does_not_apply/cannot_determine), and Missing Fact.
4. CONSERVATIVE STATUS: DO NOT mark a finding as VERIFIED if there is ANY missing project fact required to close the gap.
5. ENERGY CODE: You MUST include an Energy finding. Do not just list generic SEER2. Find the specific replacement-equipment provision in the current energy code. If unknown, set status to CONDITIONAL.
6. ACTION PLAN DEFENSIBILITY: Use "Determine whether..." instead of "Include...". Do not assume contractor licensing requirements unless explicitly found.
7. USER-PROVIDED FACTS: Explicitly label CUP/Variance as USER_PROVIDED until documents are retrieved.

OUTPUT JSON SCHEMA (Strictly follow this structure. Keep text fields concise. Do not omit fields):
{{
  "bottom_line_summary": "1-2 sentence executive summary distinguishing permit requirement from review pathway.",
  "jurisdiction": {{
    "status": "VERIFIED or CONDITIONAL",
    "county": "...",
    "city": "...",
    "building_ahj": "...",
    "planning_ahj": "...",
    "permit_portal_url": "..."
  }},
  "applicable_codes": [
    {{"code_name": "...", "edition": "...", "mandatory_date": "...", "source_url": "..."}}
  ],
  "permit_matrix": [
    {{
      "permit_type": "Mechanical",
      "permit_status": "VERIFIED", 
      "review_pathway_status": "CONDITIONAL",
      "evidence_quality": "Official AHJ checklist/application",
      "summary": "Commercial HVAC replacement requires mechanical permitting.",
      "why": "Brief explanation of the rule.",
      "evidence": "Specific source name or URL.",
      "applicability_test": {{
        "source_rule": "The specific rule found.",
        "project_fact": "The specific fact from the SOW.",
        "comparison": "How they compare.",
        "determination": "applies, does_not_apply, or cannot_determine",
        "missing_fact": "What is missing to close the gap."
      }},
      "what_i_still_need_from_you": "Specific fact the volunteer needs to provide."
    }}
  ],
  "hidden_triggers": ["List of missing details or clarifying questions."],
  "action_plan": ["Chronological step-by-step list using defensive language."]
}}

ALLOWED STATUS VALUES: "VERIFIED", "CONDITIONAL", "INFERRED", "UNKNOWN", "NOT_APPLICABLE", "USER_PROVIDED"
ALLOWED EVIDENCE QUALITY VALUES: "Direct official provision", "Official AHJ checklist/application", "Official guidance", "Secondary authoritative source", "Search result only"
"""
                result = cached_gemini_call(prompt_hash, prompt)
                st.session_state.debug_log = result.get("debug", {})
                
                if result["error"]:
                    st.error(f" {result['msg']}")
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
    
    # 1. Bottom Line Summary
    st.header("4. Research Dossier")
    st.info(f"**Bottom Line:** {data.get('bottom_line_summary', 'N/A')}")
    
    # 2. Jurisdiction
    st.subheader("📍 Jurisdiction Determination")
    jur = data.get("jurisdiction") or {}
    jur_status = jur.get('status', 'UNKNOWN')
    st.caption(f"Status: {jur_status}")
    
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**County:** {jur.get('county', 'Unknown')}")
        st.write(f"**City:** {jur.get('city', 'Unknown')}")
        st.write(f"**Building AHJ:** {jur.get('building_ahj', 'Unknown')}")
    with col2:
        st.write(f"**Planning AHJ:** {jur.get('planning_ahj', 'Unknown')}")
        portal = jur.get('permit_portal_url', '')
        if portal: st.write(f"**Portal:** [Link]({portal})")

    # 3. Codes
    st.subheader("📚 Applicable Codes & Editions")
    for code in (data.get("applicable_codes") or []):
        st.markdown(f"- **{code.get('code_name', 'Unknown')}** (Edition: {code.get('edition', 'N/A')}, Mandatory: {code.get('mandatory_date', 'N/A')})")

    # 4. Permit Matrix (Expandable with Applicability Test)
    st.subheader("📋 Permit & Review Matrix")
    
    status_map = {
        "VERIFIED": "", "INFERRED": "🟡", "CONDITIONAL": "🟠",
        "UNKNOWN": "🔴", "NOT_APPLICABLE": "⚪", "USER_PROVIDED": "🔵"
    }

    for item in data.get("permit_matrix", []):
        permit_emoji = status_map.get(item.get("permit_status", "UNKNOWN"), "")
        pathway_emoji = status_map.get(item.get("review_pathway_status", "UNKNOWN"), "")
        
        expander_title = f"{permit_emoji} {item.get('permit_type', 'Unknown')} — Permit: {item.get('permit_status', 'UNKNOWN')} | Pathway: {pathway_emoji} {item.get('review_pathway_status', 'UNKNOWN')}"
        
        with st.expander(expander_title, expanded=False):
            st.write(f"**Summary:** {item.get('summary', 'N/A')}")
            st.write(f"**Why:** {item.get('why', 'N/A')}")
            st.write(f"**Evidence Quality:** {item.get('evidence_quality', 'N/A')}")
            st.write(f"**Evidence Source:** {item.get('evidence', 'None retrieved.')}")
            
            # The Applicability Test Fields
            st.divider()
            st.subheader("Applicability Test")
            app_test = item.get("applicability_test") or {}
            st.write(f"**Source Rule:** {app_test.get('source_rule', 'N/A')}")
            st.write(f"**Project Fact:** {app_test.get('project_fact', 'N/A')}")
            st.write(f"**Comparison:** {app_test.get('comparison', 'N/A')}")
            st.write(f"**Determination:** {app_test.get('determination', 'N/A')}")
            st.write(f"**Missing Fact:** {app_test.get('missing_fact', 'N/A')}")
            
            st.divider()
            st.info(f"**❓ What I still need from you:** {item.get('what_i_still_need_from_you', 'Nothing.')}")

    # 5. Triggers & Action Plan
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Hidden Triggers & Questions")
        for trigger in (data.get("hidden_triggers") or []):
            st.markdown(f"- {trigger}")
    with col2:
        st.subheader("Volunteer Action Plan")
        for i, step in enumerate(data.get("action_plan") or [], 1):
            st.markdown(f"{i}. {step}")

    if st.session_state.sources:
        with st.expander(" Live Sources Retrieved", expanded=False):
            for i, s in enumerate(st.session_state.sources, 1):
                st.markdown(f"**{i}.** [{s['title']}]({s['url']})\n   `{s['url']}`")

    # 6. Export
    st.header("5. Export")
    col1, col2 = st.columns(2)
    
    with col1:
        doc = Document()
        doc.styles["Normal"].font.name = "Aptos"
        doc.add_heading("AHJ Research Dossier", 0)
        doc.add_paragraph(f"Project: {address} ({state})\nDate: {project_date}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
        
        doc.add_heading("Bottom Line", level=1)
        doc.add_paragraph(data.get("bottom_line_summary", ""))
        
        doc.add_heading("Jurisdiction", level=1)
        jur = data.get("jurisdiction") or {}
        doc.add_paragraph(f"Status: {jur.get('status', 'N/A')}\nCounty: {jur.get('county', 'N/A')}\nCity: {jur.get('city', 'N/A')}\nBuilding AHJ: {jur.get('building_ahj', 'N/A')}")
        
        doc.add_heading("Applicable Codes", level=1)
        for code in (data.get("applicable_codes") or []): 
            doc.add_paragraph(f"{code.get('code_name')} ({code.get('edition')}) - Mandatory: {code.get('mandatory_date')}", style='List Bullet')
        
        doc.add_heading("Permit Matrix", level=1)
        for item in data.get("permit_matrix", []):
            doc.add_heading(f"{item.get('permit_type')} - Permit: {item.get('permit_status')} | Pathway: {item.get('review_pathway_status')}", level=2)
            doc.add_paragraph(f"Summary: {item.get('summary')}")
            doc.add_paragraph(f"Why: {item.get('why')}")
            doc.add_paragraph(f"Evidence: {item.get('evidence')}")
            
            app_test = item.get("applicability_test") or {}
            doc.add_paragraph(f"Source Rule: {app_test.get('source_rule')}")
            doc.add_paragraph(f"Project Fact: {app_test.get('project_fact')}")
            doc.add_paragraph(f"Determination: {app_test.get('determination')}")
            doc.add_paragraph(f"Missing Fact: {app_test.get('missing_fact')}")
            doc.add_paragraph(f"Still needs: {item.get('what_i_still_need_from_you')}")
            
        doc.add_heading("Hidden Triggers", level=1)
        for trigger in (data.get("hidden_triggers") or []): doc.add_paragraph(trigger, style='List Bullet')
        
        doc.add_heading("Action Plan", level=1)
        for i, step in enumerate(data.get("action_plan") or [], 1): doc.add_paragraph(f"{i}. {step}")
            
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
