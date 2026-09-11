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

st.set_page_config(page_title="AHJ Research Assistant v6", page_icon="🏛️", layout="wide")

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

st.title("🏛️ AHJ Research Assistant v6")
st.caption("Discipline independence. Strict applicability testing. Vague SOW intelligence.")

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
    state = st.selectbox("State / Jurisdiction", STATE_OPTIONS, index=STATE_OPTIONS.index("Montana"))
    address = st.text_input("Project Address", "1238 MT-200, Noxon, MT")
    project_date = st.date_input("Permit / Construction Date", date.today())

with col2:
    ptype = st.selectbox("Project Type", PROJECT_TYPES, index=0)
    bclass = st.selectbox("Building / Occupancy Class", BUILDING_CLASSES, index=0)
    existing_permit = st.text_input("Existing Entitlements (Optional)", "")

# --- SECTION 2: SCOPE OF WORK ---
st.header("2. Scope of Work (SOW)")
sow_text = st.text_area(
    "Paste the complete Scope of Work below. (Even a short phrase like 'changing out existing furnace' works!)", 
    height=150,
    value="changing out existing furnace."
)

# --- SECTION 3: RESEARCH ---
st.header("3. Research Execution")

input_string = f"{state}|{address}|{project_date}|{ptype}|{bclass}|{existing_permit}|{sow_text}"
prompt_hash = hashlib.md5(input_string.encode()).hexdigest()

if st.button("🔎 Analyze & Research", type="primary", use_container_width=True):
    if mock_mode:
        st.session_state.report_data = {
            "bottom_line_summary": "Mechanical: CONDITIONAL pending AHJ confirmation. Electrical: CONDITIONAL (depends on wiring/disconnect changes). Gas/Fuel: CONDITIONAL (depends on fuel type and piping changes). Structural: NOT CURRENTLY TRIGGERED. Planning: NOT CURRENTLY TRIGGERED.",
            "immediate_questions": [
                "Is the building commercial or residential?",
                "Is the replacement furnace gas/propane, electric, or fuel oil?",
                "Is it staying in the exact same location, or being relocated?",
                "Will gas piping, electrical wiring, venting/flue, or ductwork be changed?",
                "What is the physical mounting configuration (floor, wall, rooftop)?"
            ],
            "jurisdiction": {"status": "CONDITIONAL", "county": "Sanders County", "city": "Noxon (Unincorporated)", "building_ahj": "Montana DLI (IF outside a certified local program)", "planning_ahj": "Sanders County Land Services", "permit_portal_url": "https://ebiz.mt.gov"},
            "applicable_codes": [{"code_name": "2021 International Mechanical Code (IMC)", "edition": "2021", "mandatory_date": "2022-09-01", "source_url": "https://dli.mt.gov/licensing-boards/building-codes"}],
            "permit_matrix": [
                {
                    "permit_type": "Mechanical",
                    "status": "CONDITIONAL",
                    "review_pathway_status": "CONDITIONAL",
                    "evidence_quality": "Official AHJ guidance",
                    "summary": "Permit pathway identified, subject to AHJ confirmation.",
                    "why": "If Montana DLI is the governing mechanical AHJ, commercial furnace replacement requires the applicable state mechanical permit. (Note: Building class 'Commercial' is USER-PROVIDED metadata).",
                    "evidence": "Administrative Rules of Montana (ARM) 24.301.172",
                    "what_this_does_not_establish": "Does not establish whether gas piping modifications or flue venting changes require supplementary permits.",
                    "what_i_still_need_from_you": "Confirmation of governing AHJ and furnace cut sheet showing fuel type, BTU, and venting.",
                    "applicability_test": {
                        "source_rule": "State mechanical permits are required for replacing heating systems in non-certified municipalities/counties.",
                        "project_fact": "Replacing an existing furnace. Jurisdiction is CONDITIONAL.",
                        "comparison": "Rule applies IF DLI is the AHJ, but final pathway depends on missing equipment specs.",
                        "determination": "cannot_determine",
                        "missing_fact": "Governing AHJ confirmation and equipment specification sheet."
                    },
                    "reopen_condition": "N/A"
                },
                {
                    "permit_type": "Electrical",
                    "status": "CONDITIONAL",
                    "review_pathway_status": "CONDITIONAL",
                    "evidence_quality": "Official AHJ guidance",
                    "summary": "Depends on whether electrical work is modified.",
                    "why": "ARM 24.301.401 requires an electrical permit whenever new branch circuits, altered wiring, or changed overcurrent protection devices are installed.",
                    "evidence": "ARM 24.301.401; 2020 NEC",
                    "what_this_does_not_establish": "Does not establish whether the existing circuit, breaker size, and disconnect match the new furnace.",
                    "what_i_still_need_from_you": "Existing circuit specs and new unit specs (Voltage, MCA, MOP).",
                    "applicability_test": {
                        "source_rule": "Electrical permits are required for modification of branch circuits, overcurrent protection, or disconnects.",
                        "project_fact": "SOW states 'changing out existing furnace'; electrical scope is unstated.",
                        "comparison": "If no electrical components change and new unit is compatible, no permit. If any component changes, permit required.",
                        "determination": "cannot_determine",
                        "missing_fact": "Proposed unit voltage, MCA, MOP, and whether existing wiring/disconnect will be modified."
                    },
                    "reopen_condition": "N/A"
                },
                {
                    "permit_type": "Gas / Fuel Piping",
                    "status": "CONDITIONAL",
                    "review_pathway_status": "CONDITIONAL",
                    "evidence_quality": "Official AHJ guidance",
                    "summary": "Depends on furnace fuel type and piping modifications.",
                    "why": "Fuel gas permits are required for installation, alteration, or extension of gas piping systems.",
                    "evidence": "2021 International Fuel Gas Code (IFGC)",
                    "what_this_does_not_establish": "Does not establish if the furnace is gas-fired or if existing piping is being modified.",
                    "what_i_still_need_from_you": "Confirmation of fuel type and whether gas piping is being rerouted or extended.",
                    "applicability_test": {
                        "source_rule": "Gas permits are required for modification of fuel gas piping systems.",
                        "project_fact": "Furnace fuel type and piping modification scope are unstated.",
                        "comparison": "If electric, not applicable. If gas and piping is unchanged, may not apply. If gas and piping is modified, permit required.",
                        "determination": "cannot_determine",
                        "missing_fact": "Furnace fuel type and gas piping modification scope."
                    },
                    "reopen_condition": "N/A"
                },
                {
                    "permit_type": "Planning / Zoning",
                    "status": "NOT_CURRENTLY_TRIGGERED",
                    "review_pathway_status": "SCREENING ONLY",
                    "evidence_quality": "Official guidance",
                    "summary": "No land-use trigger is identified from the current SOW.",
                    "why": "Interior mechanical replacements are typically exempt from planning review unless exterior development occurs.",
                    "evidence": "Sanders County Land Services Regulations",
                    "what_this_does_not_establish": "Does not establish parcel flood hazard status IF exterior work is later added.",
                    "what_i_still_need_from_you": "Confirmation that furnace replacement is 100% interior with no exterior equipment relocation or ground disturbance.",
                    "applicability_test": {
                        "source_rule": "Planning permits are required for exterior development, footprint expansions, or work within mapped floodplains.",
                        "project_fact": "SOW only states 'changing out existing furnace' with no mention of exterior work.",
                        "comparison": "No land-use trigger is identified from the current SOW, but missing location facts could change this.",
                        "determination": "cannot_determine",
                        "missing_fact": "Confirmation of interior-only scope vs. exterior relocation/site work."
                    },
                    "reopen_condition": "Exterior relocation, new pad, site work, or floodplain involvement is identified."
                },
                {
                    "permit_type": "Structural",
                    "status": "NOT_CURRENTLY_TRIGGERED",
                    "review_pathway_status": "SCREENING ONLY",
                    "evidence_quality": "Official guidance",
                    "summary": "No structural alteration is identified from the available SOW.",
                    "why": "Commercial mechanical equipment swaps that do not modify structural framing do not trigger structural permits.",
                    "evidence": "2021 International Building Code (IBC) Section 1613",
                    "what_this_does_not_establish": "Does not establish unit weight comparison or whether new support bracing is required.",
                    "what_i_still_need_from_you": "Unit operating weight, mounting location, and confirmation that no structural framing cuts are planned.",
                    "applicability_test": {
                        "source_rule": "Structural permits are required for structural alterations or added loads.",
                        "project_fact": "Replacing a furnace with no structural alterations specified.",
                        "comparison": "No structural trigger identified, but mounting configuration is unknown.",
                        "determination": "cannot_determine",
                        "missing_fact": "Unit weight and mounting/anchorage configuration."
                    },
                    "reopen_condition": "Furnace is rooftop, suspended, supported by modified framing, or requires structural alterations."
                },
                {
                    "permit_type": "Fire / Life Safety",
                    "status": "NOT_CURRENTLY_TRIGGERED",
                    "review_pathway_status": "SCREENING ONLY",
                    "evidence_quality": "Official guidance",
                    "summary": "No fire-code trigger is identified from the current SOW.",
                    "why": "Simple furnace replacements typically do not affect fire-rated assemblies, fire protection systems, or egress.",
                    "evidence": "2021 International Fire Code (IFC)",
                    "what_this_does_not_establish": "Does not establish if the replacement affects existing fire dampers or hazardous material storage.",
                    "what_i_still_need_from_you": "Confirmation that the replacement does not alter fire-rated assemblies or fire protection systems.",
                    "applicability_test": {
                        "source_rule": "Fire permits are required for alterations to fire protection systems, fire-rated assemblies, or hazardous materials.",
                        "project_fact": "SOW only states 'changing out existing furnace'.",
                        "comparison": "No fire-code trigger is identified from the current SOW.",
                        "determination": "does_not_apply",
                        "missing_fact": "None, unless scope changes to include fire system alterations."
                    },
                    "reopen_condition": "Scope affects fire-rated assemblies, fire protection systems, hazardous materials, or egress."
                }
            ],
            "hidden_triggers": ["If the replacement furnace requires re-routing gas supply piping, a state plumbing/fuel gas permit is triggered."],
            "action_plan": ["1. Answer the immediate questions above to clarify the scope.", "2. Collect equipment cut sheets for existing and proposed units.", "3. Verify property parcel location with County Land Services to confirm no local overlay triggers exist.", "4. Apply for applicable permits via the state or local portal."]
        }
        st.session_state.sources = [{"title": "Montana DLI Building Codes", "url": "https://dli.mt.gov/licensing-boards/building-codes"}]
        st.session_state.debug_log = {"mock": True, "note": "No API call made"}
        st.info("🛡️ Mock Mode active.")
    else:
        if not GEMINI_KEY:
            st.error("GEMINI_KEY missing.")
        else:
            with st.spinner("Analyzing scope and generating targeted research questions..."):
                prompt = f"""
You are an expert AHJ (Authority Having Jurisdiction) research assistant. You MUST output a STRICT JSON object. 

PROJECT METADATA:
- State: {state}
- Address: {address}
- Date: {project_date}
- Project Type: {ptype}
- Building Class: {bclass} (Treat this as USER-PROVIDED metadata unless verified by SOW)
- Existing Entitlements: {existing_permit}

USER-STATED SCOPE OF WORK:
{sow_text}

CRITICAL RESEARCH & APPLICABILITY RULES:
1. DISCIPLINE INDEPENDENCE: Evaluate Mechanical, Electrical, Gas/Fuel Piping, Structural, Planning/Zoning, and Fire/Life Safety as completely SEPARATE branches. Do not combine Electrical and Gas.
2. THE "NOT_APPLICABLE" BAN: You CANNOT output a discipline as "NOT_APPLICABLE" if a `missing_fact` exists that could change the outcome. Instead, you MUST use "NOT_CURRENTLY_TRIGGERED" and provide a `reopen_condition`. "NOT_APPLICABLE" is strictly reserved for things that are genuinely impossible to trigger (e.g., plumbing for a roof-only project).
3. JURISDICTION CONSISTENCY: If Jurisdiction status is "CONDITIONAL", downstream permits (like Mechanical) MUST reflect this dependency (e.g., status: "CONDITIONAL", summary: "Permit pathway identified, subject to AHJ confirmation").
4. NO SPECULATIVE BRANCHES: Do not add a Fire/Life Safety or Planning branch unless a specific trigger is in the SOW. If not, mark as "NOT_CURRENTLY_TRIGGERED" with a clear `reopen_condition`.
5. VAGUE SOW INTELLIGENCE: If the SOW is brief, prioritize generating the `immediate_questions` array with the top 5 high-value facts needed to narrow down the scope.
6. APPLICABILITY TEST: For EVERY permit type, fill out the `applicability_test` object. Use "cannot_determine" if the rule applies but a key fact is missing.

OUTPUT JSON SCHEMA (Strictly follow this structure):
{{
  "bottom_line_summary": "3-4 sentences explicitly stating the status of Mechanical, Electrical, Gas/Fuel, Planning, Structural, and Fire based on the limited SOW.",
  "immediate_questions": [
    "Top 1 high-value question",
    "Top 2 high-value question",
    "Top 3 high-value question",
    "Top 4 high-value question",
    "Top 5 high-value question"
  ],
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
      "status": "CONDITIONAL", 
      "review_pathway_status": "CONDITIONAL",
      "evidence_quality": "Official AHJ guidance",
      "summary": "Brief summary of the finding.",
      "why": "Brief explanation. Explicitly mention if relying on USER-PROVIDED metadata or conditional jurisdiction.",
      "evidence": "Specific source name or URL.",
      "what_this_does_not_establish": "Crucial: What gap remains?",
      "what_i_still_need_from_you": "Crucial: What specific fact is needed?",
      "applicability_test": {{
        "source_rule": "The specific rule found.",
        "project_fact": "The specific fact from the SOW or metadata.",
        "comparison": "How they compare.",
        "determination": "applies, does_not_apply, or cannot_determine",
        "missing_fact": "What is missing to close the gap."
      }},
      "reopen_condition": "N/A, OR a specific condition that would trigger this branch (e.g., 'Exterior relocation or site work is identified')."
    }}
  ],
  "hidden_triggers": ["List of missing details or clarifying questions."],
  "action_plan": ["Chronological step-by-step list."]
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
    st.info(f"**Bottom Line:** {data.get('bottom_line_summary', 'N/A')}")

    # 2. Immediate Questions
    questions = data.get("immediate_questions") or []
    if questions:
        st.subheader("❓ Immediate Questions (Information Needed)")
        st.warning("The current Scope of Work is brief. Before definitive permit pathways can be established, please clarify these high-value facts:")
        for i, q in enumerate(questions, 1):
            st.markdown(f"{i}. **{q}**")

    # 3. Jurisdiction
    st.subheader("📍 Jurisdiction Determination")
    jur = data.get("jurisdiction") or {}
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**County:** {jur.get('county', 'Unknown')}")
        st.write(f"**City:** {jur.get('city', 'Unknown')}")
        st.write(f"**Building AHJ:** {jur.get('building_ahj', 'Unknown')}")
    with col2:
        st.write(f"**Planning AHJ:** {jur.get('planning_ahj', 'Unknown')}")
        portal = jur.get('permit_portal_url', '')
        if portal: st.write(f"**Portal:** [Link]({portal})")

    # 4. Codes
    st.subheader("📚 Applicable Codes & Editions")
    for code in (data.get("applicable_codes") or []):
        st.markdown(f"- **{code.get('code_name', 'Unknown')}** (Edition: {code.get('edition', 'N/A')}, Mandatory: {code.get('mandatory_date', 'N/A')})")

    # 5. Permit Matrix
    st.subheader("📋 Permit & Review Matrix")
    
    status_map = {
        "VERIFIED": "", "INFERRED": "🟡", "CONDITIONAL": "🟠",
        "UNKNOWN": "🔴", "NOT_APPLICABLE": "⚪", "NOT_CURRENTLY_TRIGGERED": "⚪", "USER_PROVIDED": ""
    }

    for item in data.get("permit_matrix", []):
        status_emoji = status_map.get(item.get("status", "UNKNOWN"), "")
        pathway_emoji = status_map.get(item.get("review_pathway_status", "UNKNOWN"), "")
        expander_title = f"{status_emoji} {item.get('permit_type', 'Unknown')} — Permit: {item.get('status', 'UNKNOWN')} | Pathway: {pathway_emoji} {item.get('review_pathway_status', 'UNKNOWN')}"
        
        with st.expander(expander_title, expanded=False):
            st.write(f"**Summary:** {item.get('summary', 'N/A')}")
            st.write(f"**Why:** {item.get('why', 'N/A')}")
            st.write(f"**Evidence Quality:** {item.get('evidence_quality', 'N/A')}")
            st.write(f"**Evidence Source:** {item.get('evidence', 'None retrieved.')}")
            
            st.divider()
            st.warning(f"**⚠️ What this does NOT establish:** {item.get('what_this_does_not_establish', 'Nothing.')}")
            st.info(f"**❓ What I still need from you:** {item.get('what_i_still_need_from_you', 'Nothing.')}")
            
            reopen = item.get('reopen_condition', 'N/A')
            if reopen and reopen.lower() != 'n/a':
                st.success(f"**🔄 Reopen if:** {reopen}")
            
            st.subheader("Applicability Test Logic")
            app_test = item.get("applicability_test") or {}
            st.write(f"**Source Rule:** {app_test.get('source_rule', 'N/A')}")
            st.write(f"**Project Fact:** {app_test.get('project_fact', 'N/A')}")
            st.write(f"**Comparison:** {app_test.get('comparison', 'N/A')}")
            st.write(f"**Determination:** {app_test.get('determination', 'N/A').replace('_', ' ').title()}")
            st.write(f"**Missing Fact:** {app_test.get('missing_fact', 'N/A')}")

    # 6. Triggers & Action Plan
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Hidden Triggers & Questions")
        for trigger in (data.get("hidden_triggers") or []):
            st.markdown(f"- {trigger}")
    with col2:
        st.subheader("Volunteer Action Plan")
        for i, step in enumerate(data.get("action_plan") or [], 1):
            clean_step = re.sub(r'^\d+\.\s*', '', step).strip()
            st.markdown(f"{i}. {clean_step}")

    if st.session_state.sources:
        with st.expander("🔗 Live Sources Retrieved", expanded=False):
            for i, s in enumerate(st.session_state.sources, 1):
                st.markdown(f"**{i}.** [{s['title']}]({s['url']})\n   `{s['url']}`")

    # 7. Export
    st.header("5. Export")
    col1, col2 = st.columns(2)
    
    with col1:
        doc = Document()
        doc.styles["Normal"].font.name = "Aptos"
        doc.add_heading("AHJ Research Dossier", 0)
        doc.add_paragraph(f"Project: {address} ({state})\nDate: {project_date}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
        
        doc.add_heading("Bottom Line", level=1)
        doc.add_paragraph(data.get("bottom_line_summary", ""))
        
        doc.add_heading("Immediate Questions", level=1)
        for i, q in enumerate(data.get("immediate_questions") or [], 1):
            doc.add_paragraph(f"{i}. {q}")
        
        doc.add_heading("Jurisdiction", level=1)
        jur = data.get("jurisdiction") or {}
        doc.add_paragraph(f"Status: {jur.get('status', 'N/A')}\nCounty: {jur.get('county', 'N/A')}\nCity: {jur.get('city', 'N/A')}\nBuilding AHJ: {jur.get('building_ahj', 'N/A')}")
        
        doc.add_heading("Applicable Codes", level=1)
        for code in (data.get("applicable_codes") or []): 
            doc.add_paragraph(f"{code.get('code_name')} ({code.get('edition')}) - Mandatory: {code.get('mandatory_date')}", style='List Bullet')
        
        doc.add_heading("Permit Matrix", level=1)
        for item in data.get("permit_matrix", []):
            doc.add_heading(f"{item.get('permit_type')} - Permit: {item.get('status')} | Pathway: {item.get('review_pathway_status')}", level=2)
            doc.add_paragraph(f"Summary: {item.get('summary')}")
            doc.add_paragraph(f"Why: {item.get('why')}")
            doc.add_paragraph(f"Evidence: {item.get('evidence')}")
            doc.add_paragraph(f"Does NOT establish: {item.get('what_this_does_not_establish')}")
            doc.add_paragraph(f"Still needs: {item.get('what_i_still_need_from_you')}")
            
            reopen = item.get('reopen_condition', 'N/A')
            if reopen and reopen.lower() != 'n/a':
                doc.add_paragraph(f"Reopen if: {reopen}")
            
            app_test = item.get("applicability_test") or {}
            doc.add_paragraph(f"Source Rule: {app_test.get('source_rule')}")
            doc.add_paragraph(f"Project Fact: {app_test.get('project_fact')}")
            doc.add_paragraph(f"Determination: {app_test.get('determination', '').replace('_', ' ').title()}")
            doc.add_paragraph(f"Missing Fact: {app_test.get('missing_fact')}")
            
        doc.add_heading("Hidden Triggers", level=1)
        for trigger in (data.get("hidden_triggers") or []): doc.add_paragraph(trigger, style='List Bullet')
        
        doc.add_heading("Action Plan", level=1)
        for i, step in enumerate(data.get("action_plan") or [], 1): 
            clean_step = re.sub(r'^\d+\.\s*', '', step).strip()
            doc.add_paragraph(f"{i}. {clean_step}")
            
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
