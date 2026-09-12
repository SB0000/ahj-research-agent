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

st.set_page_config(page_title="AHJ Research Assistant v21", page_icon="🏛️", layout="wide")

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
PROMPT_VERSION = "v21_generalized_reasoning_engine"

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

    # Collect evidence IDs and build map
    evidence_items = data.get("evidence", [])
    evidence_by_id = {
        item.get("id"): item
        for item in evidence_items
        if item.get("id")
    }
    evidence_ids = set(evidence_by_id.keys())

    # Validate evidence records
    for evidence in evidence_items:
        eid = evidence.get("id", "Unknown")
        if not evidence.get("title"): errors.append(f"{eid}: evidence missing title.")
        if not evidence.get("url"): errors.append(f"{eid}: evidence missing URL.")
        if not evidence.get("authority"): errors.append(f"{eid}: evidence missing authority.")
        if not evidence.get("discipline"): errors.append(f"{eid}: evidence missing discipline.")
        if not evidence.get("rule"): errors.append(f"{eid}: evidence missing rule.")

    # Validate jurisdiction evidence
    jurisdiction = data.get("jurisdiction", {})
    for eid in jurisdiction.get("evidence", []):
        if eid not in evidence_ids:
            errors.append(f"Jurisdiction references nonexistent evidence ID: {eid}")
    if jurisdiction.get("status") == "VERIFIED" and not jurisdiction.get("evidence"):
        errors.append("Verified jurisdiction must have evidence.")

    # Validate code evidence
    for code in data.get("codes", []):
        for eid in code.get("evidence", []):
            if eid not in evidence_ids:
                errors.append(f"Code '{code.get('name')}' references nonexistent evidence ID: {eid}")
        if code.get("status") == "CURRENT" and not code.get("evidence"):
            errors.append(f"Current code '{code.get('name')}' must have evidence.")

    # Validate disciplines
    for item in data.get("disciplines", []):
        discipline = item.get("type", "Unknown")
        permit = item.get("permit")
        pathway = item.get("pathway")
        app = item.get("applicability", {})
        determination = app.get("determination")
        relationship = app.get("relationship")

        if permit not in allowed_statuses: errors.append(f"{discipline}: invalid permit status '{permit}'")
        if pathway not in allowed_statuses: errors.append(f"{discipline}: invalid pathway status '{pathway}'")
        if determination not in allowed_determinations: errors.append(f"{discipline}: invalid determination '{determination}'")
        if relationship not in allowed_relationships: errors.append(f"{discipline}: invalid relationship '{relationship}'")

        # Evidence references must exist
        for eid in item.get("evidence", []):
            if eid not in evidence_ids:
                errors.append(f"{discipline}: references nonexistent evidence ID '{eid}'")
            
            # Cross-discipline evidence check
            evidence = evidence_by_id.get(eid)
            if evidence:
                ev_disc = (evidence.get("discipline") or "").lower()
                if ev_disc and discipline.lower() and ev_disc != discipline.lower() and relationship != "direct":
                    errors.append(f"{discipline}: evidence {eid} belongs to {ev_disc} but relationship is not 'direct'.")

        # VERIFIED_REQUIRED requires evidence
        if permit == "VERIFIED_REQUIRED" and not item.get("evidence"):
            errors.append(f"{discipline}: VERIFIED_REQUIRED requires evidence.")

        # Direct relationship requires an actual rule
        if relationship == "direct" and not app.get("rule"):
            errors.append(f"{discipline}: direct relationship requires a source rule.")

        # cannot_determine requires missing information
        if determination == "cannot_determine" and not app.get("missing"):
            errors.append(f"{discipline}: cannot_determine requires missing facts.")

        # Missing facts cannot coexist with definitive does_not_apply
        if app.get("missing") and determination == "does_not_apply":
            errors.append(f"{discipline}: unresolved facts are present, but determination is does_not_apply.")

        # NOT_CURRENTLY_TRIGGERED should not have definitive does_not_apply when unresolved facts remain
        if permit == "NOT_CURRENTLY_TRIGGERED" and determination == "does_not_apply" and app.get("missing"):
            errors.append(f"{discipline}: NOT_CURRENTLY_TRIGGERED cannot be paired with does_not_apply when missing facts could change applicability.")

        # NOT_APPLICABLE requires defensible determination
        if permit == "NOT_APPLICABLE" and determination != "does_not_apply":
            errors.append(f"{discipline}: NOT_APPLICABLE should have determination=does_not_apply.")

        # not_established relationship cannot claim a verified downstream requirement
        if relationship == "not_established":
            finding = (item.get("permit_finding") or item.get("finding", "")).lower()
            dangerous_phrases = ["is required", "requires", "shall require", "automatically triggers", "therefore requires", "must obtain"]
            for phrase in dangerous_phrases:
                if phrase in finding:
                    errors.append(f"{discipline}: relationship is not_established but finding claims a downstream requirement: '{phrase}'.")
                    break

        # Specific threshold/pathway claims should have evidence
        finding = (item.get("permit_finding") or item.get("finding", "")).lower()
        threshold_markers = ["lb", "lbs", "cfm", "ton", "tons", "btu", "square feet", "sq ft", "feet", "foot", "percent", "%", "section", "chapter", "threshold"]
        if any(marker in finding for marker in threshold_markers) and not item.get("evidence"):
            errors.append(f"{discipline}: finding contains a threshold/code claim but has no evidence.")

        # Pathway cannot be more certain than the evidence
        if pathway == "VERIFIED_REQUIRED":
            if not item.get("evidence"):
                errors.append(f"{discipline}: verified pathway requires evidence.")
            if relationship == "not_established":
                errors.append(f"{discipline}: pathway cannot be verified when relationship is not_established.")

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

st.title("🏛️ AHJ Research Assistant v21")
st.caption("Generalized reasoning engine. Applicability vs. Compliance. Cross-discipline firewall.")

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
            "bottom_line": "Mechanical permit VERIFIED REQUIRED. Energy code applies but compliance is unknown. Electrical and Planning are CONDITIONAL pending missing facts.",
            "jurisdiction": {"status": "CONDITIONAL", "county": "Washington County", "city": "Hillsboro", "ahj": "Unresolved - boundary unconfirmed", "evidence": ["E1"]},
            "codes": [{"name": "2025 Oregon Mechanical Specialty Code", "status": "CURRENT", "evidence": ["E2"]}],
            "evidence": [
                {"id": "E1", "title": "Washington County Building Services", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Jurisdiction", "rule": "Jurisdiction for unincorporated areas."},
                {"id": "E2", "title": "Washington County Mechanical Unit Checklist", "url": "https://www.washingtoncounty.org/1134/Building-Services", "authority": "county", "discipline": "Mechanical", "rule": "Commercial mechanical permit required for replacement."},
                {"id": "E3", "title": "Oregon Energy Efficiency Specialty Code", "url": "https://www.oregon.gov/bcd", "authority": "state", "discipline": "Energy", "rule": "Replacement mechanical equipment must comply with current energy efficiency standards."}
            ],
            "disciplines": [
                {
                    "type": "Mechanical",
                    "applicability": {"rule": "Permit required for regulated mechanical replacement.", "fact": "Commercial HVAC replacement.", "determination": "applies", "missing": "", "relationship": "direct"},
                    "permit": "VERIFIED_REQUIRED",
                    "permit_finding": "Commercial mechanical permit required for replacement.",
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Pathway depends on unit specs (weight/CFM) to determine if minor exemption applies.",
                    "evidence": ["E2"],
                    "missing": ["Proposed unit weight and CFM for pathway determination."],
                    "reopen": []
                },
                {
                    "type": "Electrical",
                    "applicability": {"rule": "Permits required for branch circuit/disconnect modification.", "fact": "SOW states different brand, electrical scope unstated.", "determination": "cannot_determine", "missing": "Unit electrical specs and scope of electrical changes.", "relationship": "conditional"},
                    "permit": "CONDITIONAL",
                    "permit_finding": "Permit required only if electrical work is modified.",
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Pathway cannot be determined until electrical scope is defined.",
                    "evidence": ["E1"],
                    "missing": ["Unit electrical specs (MCA, MOP, voltage)", "Scope of electrical changes."],
                    "reopen": ["Modifying electrical disconnect, wiring, or breaker."]
                },
                {
                    "type": "Energy",
                    "applicability": {"rule": "Current energy code applies to replacement equipment.", "fact": "Replacing HVAC unit; efficiency ratings unknown.", "determination": "applies", "missing": "", "relationship": "direct"},
                    "permit": "CONDITIONAL",
                    "permit_finding": "Energy code compliance applies, but separate permit requirement is unconfirmed.",
                    "pathway": "CONDITIONAL",
                    "pathway_finding": "Compliance pathway depends on equipment specifications.",
                    "evidence": ["E3"],
                    "missing": ["Replacement equipment efficiency/specifications to prove compliance."],
                    "reopen": []
                },
                {
                    "type": "Structural",
                    "applicability": {"rule": "Structural permits required for alterations or added loads.", "fact": "Ground-level replacement on existing pad; no structural work stated.", "determination": "cannot_determine", "missing": "Replacement unit operating weight and anchorage configuration.", "relationship": "not_established"},
                    "permit": "NOT_CURRENTLY_TRIGGERED",
                    "permit_finding": "No structural alteration identified from current scope.",
                    "pathway": "NOT_CURRENTLY_TRIGGERED",
                    "pathway_finding": "No structural review pathway triggered.",
                    "evidence": [],
                    "missing": ["Replacement unit operating weight", "Anchorage configuration."],
                    "reopen": ["Rooftop mounting, suspended installation, or structural framing modifications occur."]
                },
                {
                    "type": "Planning / CUP",
                    "applicability": {"rule": "Work must comply with existing CUP conditions.", "fact": "Parcel operates under existing CUP; conditions not retrieved.", "determination": "cannot_determine", "missing": "Actual CUP conditions governing exterior equipment.", "relationship": "conditional"},
                    "permit": "NOT_CURRENTLY_TRIGGERED",
                    "permit_finding": "No new land-use trigger identified from current scope.",
                    "pathway": "NOT_CURRENTLY_TRIGGERED",
                    "pathway_finding": "No planning review pathway triggered unless CUP conditions are violated.",
                    "evidence": ["E1"],
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

APPLICABILITY VS COMPLIANCE
Do not confuse whether a regulation applies with whether the project has demonstrated compliance.
- The energy code may APPLY.
- Equipment efficiency information may be UNKNOWN.
- A separate energy permit may NOT be established.
Therefore these may legitimately coexist: applicability="applies", permit="CONDITIONAL", missing="equipment compliance information".
Do not mark a regulation as NOT_APPLICABLE merely because project compliance information is missing.

UNRESOLVED FACT RULE
If an unresolved project fact could change whether a regulation applies, use determination="cannot_determine". Do not use "does_not_apply" when the missing fact could change the conclusion. However, if the rule's applicability is already established and the missing information only affects compliance, permit pathway, or documentation, the applicability determination may remain "applies" while permit/pathway remain CONDITIONAL or UNKNOWN.

EVIDENCE CHAIN RULE
Never infer a downstream regulatory consequence from a threshold, exemption, definition, scope rule, or general requirement unless authoritative evidence explicitly establishes that consequence.
Example: If a source establishes "Equipment over X is outside exemption Y", you may conclude "Exemption Y does not apply if the equipment exceeds X." You may NOT automatically conclude "Structural engineering is required" unless authoritative structural evidence establishes that relationship.

DISCIPLINE INDEPENDENCE RULE
Evidence discovered for Discipline A cannot establish a regulatory requirement for Discipline B unless the authoritative source explicitly connects the two. When a cross-discipline consequence is suspected but not explicitly established, mark the relationship "not_established" and explain what authoritative evidence would be required.

BOTTOM LINE RULE
The bottom line must be a synthesis of the discipline findings. Do not introduce a conclusion in the bottom line that does not appear in the underlying evidence-backed discipline findings. Never upgrade CONDITIONAL, UNKNOWN, or CANNOT_DETERMINE into a definitive requirement in the bottom line.

CRITICAL RULES:
1. NO ARTIFICIAL CAPS: Do not limit the number of disciplines or evidence items.
2. JURISDICTION: Establish the actual building/AHJ jurisdiction. If unclear, set status="CONDITIONAL" and ahj="Unresolved - boundary unconfirmed".
3. FEDERAL/TRIBAL/HISTORIC: Actively check for federal waterways, tribal land, or historic districts. Add a specific discipline if applicable.
4. STATUS VALUES: VERIFIED_REQUIRED, CONDITIONAL, INFERRED, UNKNOWN, NOT_APPLICABLE, NOT_CURRENTLY_TRIGGERED, USER_PROVIDED.
5. DETERMINATION VALUES: applies, does_not_apply, cannot_determine.
6. RELATIONSHIP VALUES: direct, conditional, not_established.
7. DO NOT SPECULATE: Never invent dates, thresholds, or requirements without retrieved evidence.

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
        "fact": "What is known about this project.",
        "determination": "applies|does_not_apply|cannot_determine",
        "missing": "What fact prevents a final applicability determination, if any.",
        "relationship": "direct|conditional|not_established"
      }},
      "permit": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
      "permit_finding": "What the evidence establishes about the permit requirement.",
      "pathway": "VERIFIED_REQUIRED|CONDITIONAL|INFERRED|UNKNOWN|NOT_APPLICABLE|NOT_CURRENTLY_TRIGGERED|USER_PROVIDED",
      "pathway_finding": "What the evidence establishes about the review/submittal pathway.",
      "evidence": ["E1"],
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
            st.markdown(f"**Applicability:** {app.get('determination', 'N/A').replace('_', ' ').title()}")
            st.markdown(f"**Source Rule:** {app.get('rule', 'N/A')}")
            st.markdown(f"**Project Fact:** {app.get('fact', 'N/A')}")
            st.markdown(f"**Relationship:** {app.get('relationship', 'N/A').replace('_', ' ').title()}")
            
            st.divider()
            st.markdown(f"**Permit:** {item.get('permit', 'N/A')}")
            st.markdown(f"**Permit Finding:** {item.get('permit_finding', 'N/A')}")
            
            st.markdown(f"**Pathway:** {item.get('pathway', 'N/A')}")
            st.markdown(f"**Pathway Finding:** {item.get('pathway_finding', 'N/A')}")
            
            st.divider()
            ev_ids = item.get("evidence", [])
            if ev_ids:
                st.write("**Evidence:**")
                for eid in ev_ids:
                    if eid in ev_dict:
                        ev = ev_dict[eid]
                        st.markdown(f"- **[{ev['title']}]({ev['url']})** `[{ev.get('authority', 'other').upper()}]` `[{ev.get('discipline', 'general').upper()}]`")
                        st.caption(f"  *Rule:* {ev.get('rule', 'N/A')}")
            else:
                st.write("**Evidence:** None retrieved.")
            
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
            doc.add_paragraph(f"Applicability: {app.get('determination', '').replace('_', ' ').title()}")
            doc.add_paragraph(f"Source Rule: {app.get('rule')}")
            doc.add_paragraph(f"Project Fact: {app.get('fact')}")
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
