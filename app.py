import os
import re
import json
import time
from datetime import datetime, date
from io import BytesIO

import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
import urllib.request
import urllib.error

# ============================================================
# AHJ RESEARCH ASSISTANT (Quota-Safe, Single-Pass Version)
# ============================================================

st.set_page_config(page_title="AHJ Research Assistant", page_icon="🏛️", layout="wide")

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
    "HVAC Replacement", "Reroof", "Parking / Site Improvements",
    "Interior Remodel", "Major Remodel", "Addition", "New Construction", "Other",
]

BUILDING_CLASSIFICATIONS = [
    "Commercial", "Assembly", "Institutional", "Industrial",
    "Agricultural", "Residential", "Mixed-use", "Unknown",
]

CONFIDENCE_GUIDE = """
🟢 VERIFIED: A specific authoritative source was retrieved and supports the statement.
🟡 LIKELY: Strong preliminary conclusion, but exact project/AHJ applicability needs confirmation.
🟠 CONDITIONAL: True only if a stated condition applies.
🔴 UNKNOWN: Available evidence does not establish the answer.
⚠️ VERIFY: A direct AHJ, parcel record, permit record, or code official check is required.
"""

# -----------------------------
# Session State
# -----------------------------
defaults = {
    "report": "", "research_sources": [], "verified_facts": [], 
    "source_checks": [], "project_fingerprint": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# -----------------------------
# Secrets / Model Configuration
# -----------------------------
def secret_or_env(name, default=""):
    try:
        value = st.secrets.get(name, None)
        if value: return value
    except Exception:
        pass
    return os.getenv(name, default)

GEMINI_KEY = secret_or_env("GEMINI_KEY", "")
MODEL_NAME = secret_or_env("GEMINI_MODEL", "gemini-3.6-flash")

# -----------------------------
# Helpers
# -----------------------------
def extract_urls(text):
    if not text: return []
    return list(dict.fromkeys(url.rstrip(".,;:") for url in re.findall(r"https?://[^\s\]\)>\"']+", text)))

def check_url(url, timeout=5):
    result = {"url": url, "reachable": False, "status": "", "note": ""}
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "AHJ-Research/3.0"})
        with urllib.request.urlopen(req, timeout=timeout) as response:
            result["reachable"] = 200 <= response.status < 400
            result["status"] = str(response.status)
    except urllib.error.HTTPError as exc:
        result["status"] = str(exc.code)
    except Exception as exc:
        result["note"] = type(exc).__name__
    return result

def build_fingerprint(state, project_date, project_type, classification, address, details, scope_answers):
    fp = {
        "state": state, "project_date": project_date.isoformat() if isinstance(project_date, date) else str(project_date),
        "project_type": project_type, "building_classification": classification,
        "address": address.strip(), "details": details.strip(),
    }
    fp.update(scope_answers)
    return fp

def compact_fingerprint(fp):
    return "\n".join(f"- {k.replace('_', ' ').title()}: {v}" for k, v in fp.items() if v not in ("", None, "No answer", []))

def base_system_rules():
    return """
You are an AHJ (Authority Having Jurisdiction) research assistant for construction volunteers.
NON-NEGOTIABLE RULES:
1. LIVE SEARCH REQUIRED: You MUST use your live Google Search tool to find current code editions, effective dates, and official government sources. Do not rely on internal training memory.
2. OFFICIAL SOURCES FIRST: Prefer State building-code agencies, County/City/AHJ portals, and Official fire authorities.
3. NO INVENTED DATA: Never invent section numbers, permit thresholds, or URLs. If you cannot verify it via search, state UNKNOWN or VERIFY.
4. CONCISE OUTPUT: Volunteers need to understand the result in under two minutes. Be highly concise.
"""

# -----------------------------
# QUOTA-SAFE GEMINI CALL WITH AUTO-RETRY
# -----------------------------
def call_gemini(prompt, max_retries=3):
    if not GEMINI_KEY:
        return {"text": "ERROR: GEMINI_KEY is not configured.", "sources": [], "error": True}

    for attempt in range(max_retries):
        try:
            client = genai.Client(api_key=GEMINI_KEY)
            
            # Strict token limit to prevent quota exhaustion
            config = types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
                max_output_tokens=2500, 
            )

            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config=config,
            )

            text = getattr(response, "text", "").strip()
            sources = []

            # Extract live sources from grounding metadata
            try:
                candidates = getattr(response, "candidates", [])
                if candidates:
                    metadata = getattr(candidates[0], "grounding_metadata", None)
                    if metadata and getattr(metadata, "grounding_chunks", None):
                        for chunk in metadata.grounding_chunks:
                            if getattr(chunk, "web", None):
                                sources.append({
                                    "title": getattr(chunk.web, "title", "Source"),
                                    "url": getattr(chunk.web, "uri", "")
                                })
            except Exception:
                pass

            return {"text": text if text else "ERROR: No text output generated.", "sources": sources, "error": False}

        except Exception as exc:
            error_str = str(exc)
            if "429" in error_str and attempt < max_retries - 1:
                wait_time = 2 ** attempt  # Exponential backoff: 1s, 2s, 4s
                time.sleep(wait_time)
                continue  # Retry
            return {"text": f"ERROR: API failed after {max_retries} attempts: {error_str}", "sources": [], "error": True}

# -----------------------------
# SINGLE-PASS MASTER PROMPT
# -----------------------------
def master_research_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT DETAILS:
{compact_fingerprint(fp)}

TASK: Perform a comprehensive, single-pass research analysis for this project. 
You MUST use your live Google Search tool to find current, official sources for jurisdiction, code editions, and permit requirements.

Think step-by-step internally, then output ONLY the final report in this EXACT format:

# PRELIMINARY ANSWER
(Brief practical answer)
| Issue | Result | Confidence |
|---|---|---|

# CURRENT CODE PATH
(Applicable code families, editions, effective dates, with [Source] citations)

# WHAT WE KNOW
(Max 6 bullets, cited)

# WHAT COULD CHANGE THE ANSWER
(Max 4 bullets)

# QUESTIONS FOR THE AHJ
(Max 4 precise questions)

# VOLUNTEER ACTION LIST
(Max 5 steps)

# SOURCES
(List of URLs found during search)

# RESEARCH LIMITATIONS
(1 short paragraph on what could not be independently established)
"""

def clean_report(text):
    if not text: return text
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def source_diagnostics(report, sources):
    urls = []
    for source in sources:
        url = source.get("url")
        if url and url not in urls: urls.append(url)
    for url in extract_urls(report):
        if url not in urls: urls.append(url)
    return [check_url(url) for url in urls[:20]]

def build_docx(fp, report, verified_facts, diagnostics, sources):
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)

    doc.add_heading("AHJ Research Report", 0)
    doc.add_paragraph(f"Project: {fp.get('project_type', '')}\nState: {fp.get('state', '')}\nAddress: {fp.get('address', '')}\nDate: {fp.get('project_date', '')}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
    doc.add_paragraph("DRAFT RESEARCH AID — Verify applicable requirements with the AHJ before relying on this report.")

    doc.add_heading("Project Scope", level=1)
    doc.add_paragraph(compact_fingerprint(fp))
    doc.add_heading("Research Report", level=1)
    
    for line in report.splitlines():
        line = line.strip()
        if not line: continue
        if line.startswith("# "): doc.add_heading(line[2:], level=1)
        elif line.startswith("## "): doc.add_heading(line[3:], level=2)
        elif line.startswith("- "): doc.add_paragraph(line[2:], style="List Bullet")
        elif line.startswith("|"): doc.add_paragraph(line)
        else: doc.add_paragraph(line)

    if sources:
        doc.add_heading("Live Sources Retrieved", level=1)
        for i, source in enumerate(sources, 1):
            doc.add_paragraph(f"[S{i}] {source.get('title', 'Source')}\n{source.get('url', '')}", style="List Bullet")

    if verified_facts:
        doc.add_heading("Volunteer-Verified Facts", level=1)
        for fact in verified_facts:
            doc.add_paragraph(f"{fact.get('fact', '')}\nSource: {fact.get('source', '')}", style="List Bullet")

    if diagnostics:
        doc.add_heading("Source URL Diagnostics", level=1)
        for item in diagnostics:
            status = "Reachable" if item["reachable"] else "Not verified"
            doc.add_paragraph(f"{status}: {item['url']} ({item.get('status', '')})")

    doc.add_heading("Confidence Guide", level=1)
    doc.add_paragraph(CONFIDENCE_GUIDE)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ============================================================
# UI
# ============================================================
st.title("🏛️ AHJ Research Assistant")
st.caption("Quota-safe, search-grounded research. Official sources first. Human verification for final decisions.")

with st.sidebar:
    st.info("This app uses a single, highly optimized API call with native Google Search to stay well within free-tier limits while delivering live, cited results.")
    st.divider()
    st.subheader("Confidence Guide")
    st.markdown(CONFIDENCE_GUIDE)

st.header("1. Project")
col1, col2 = st.columns(2)

with col1:
    state = st.selectbox("State / jurisdiction", STATE_OPTIONS, index=STATE_OPTIONS.index("Oregon"))
    address = st.text_input("Project address", value="24340 NW Meek Rd, 97124")
    project_date = st.date_input("Project / permit date", value=date.today())
    project_type = st.selectbox("Project type", PROJECT_TYPES, index=0)

with col2:
    classification = st.selectbox("Building / use classification", BUILDING_CLASSIFICATIONS, index=1)
    existing_permit = st.text_input("Existing permit / land-use condition", value="Existing Conditional Use Permit (CUP)")
    building_status = st.selectbox("Building status", ["Existing building", "Existing building — alteration/remodel", "New construction", "Addition", "Unknown"], index=0)

details = st.text_area("Project description", value="Ground-level exterior HVAC unit replacement; like-for-like replacement with a different brand on an existing exterior pad.", height=100)

st.header("2. Scope Details")
scope = {"building_status": building_status}

if project_type == "HVAC Replacement":
    a, b = st.columns(2)
    with a:
        scope["equipment_location"] = st.selectbox("Equipment location", ["Ground-level exterior", "Rooftop", "Interior", "Other"], index=0)
        scope["same_location"] = st.selectbox("Same exact location?", ["Yes", "No", "Unknown"], index=0)
        scope["existing_pad"] = st.selectbox("Existing pad/slab reused?", ["Yes", "No", "Unknown"], index=0)
        scope["ductwork"] = st.selectbox("Ductwork affected?", ["No", "Yes", "Unknown"], index=0)
    with b:
        scope["roof_penetrations"] = st.selectbox("Roof penetrations?", ["No", "Yes", "Unknown"], index=0)
        scope["site_disturbance"] = st.selectbox("New site disturbance?", ["No", "Yes", "Unknown"], index=0)
        scope["electrical_changes"] = st.selectbox("Electrical changes?", ["Unknown", "No — same circuit/disconnect", "Yes — breaker/wiring/disconnect changes"], index=0)
        scope["refrigerant"] = st.selectbox("Replacement refrigerant known?", ["Unknown", "A2L / R-32 / R-454B", "Non-A2L", "Other"], index=0)
elif project_type == "Reroof":
    scope["roof_type"] = st.selectbox("Roof type", ["Unknown", "Low-slope commercial", "Steep-slope", "Other"])
    scope["tear_off"] = st.selectbox("Tear-off or recover?", ["Unknown", "Tear-off", "Recover / overlay"])
    scope["equipment_affected"] = st.selectbox("Rooftop equipment affected?", ["No", "Yes", "Unknown"])
else:
    scope["scope_details"] = st.text_area("Additional scope details", height=100)

scope["existing_land_use"] = existing_permit
fp = build_fingerprint(state, project_date, project_type, classification, address, details, scope)
st.session_state["project_fingerprint"] = fp

with st.expander("See the scope fingerprint the research engine will use"):
    st.code(compact_fingerprint(fp))

st.header("3. Research")
run_research = st.button("🔎 Research Project", type="primary", use_container_width=True)

if run_research:
    if not address.strip():
        st.error("Enter a project address.")
    elif not GEMINI_KEY:
        st.error("GEMINI_KEY is missing. Add it to Streamlit secrets or environment variables.")
    else:
        with st.spinner("Searching live web for official sources, code editions, and AHJ requirements (this may take 10-15 seconds)..."):
            prompt = master_research_prompt(fp)
            result = call_gemini(prompt)
            
            st.session_state["research_sources"] = result.get("sources", [])
            st.session_state["report"] = clean_report(result["text"])
            
            if not result["error"]:
                with st.spinner("Checking source URLs for basic reachability..."):
                    st.session_state["source_checks"] = source_diagnostics(st.session_state["report"], st.session_state["research_sources"])
                st.success(f"Research complete. {len(st.session_state['research_sources'])} live source(s) retrieved.")
            else:
                st.error(result["text"])

# -----------------------------
# Results
# -----------------------------
if st.session_state["report"] and not st.session_state["report"].startswith("ERROR"):
    st.header("4. Preliminary Result")
    st.markdown(st.session_state["report"])

    if st.session_state["research_sources"]:
        with st.expander("Sources actually retrieved by Gemini"):
            for i, source in enumerate(st.session_state["research_sources"], 1):
                st.markdown(f"**[S{i}] {source.get('title', 'Source')}**  \n{source.get('url', '')}")

    if st.session_state["source_checks"]:
        with st.expander("Source URL diagnostics"):
            for item in st.session_state["source_checks"]:
                if item["reachable"]:
                    st.success(f"Reachable: {item['url']} ({item.get('status', '')})")
                else:
                    st.warning(f"Not verified: {item['url']}")

    st.header("5. Volunteer Verification")
    with st.form("verification_form"):
        verified_fact = st.text_input("Verified fact")
        verified_source = st.text_input("Source / person / record")
        if st.form_submit_button("Add verified fact") and verified_fact.strip():
            st.session_state["verified_facts"].append({"fact": verified_fact.strip(), "source": verified_source.strip(), "date": datetime.now().isoformat()})
            st.success("Added to verification log.")

    if st.session_state["verified_facts"]:
        for i, fact in enumerate(st.session_state["verified_facts"], 1):
            st.markdown(f"**{i}. 🟢 VERIFIED** — {fact['fact']}  \nSource: {fact['source']}")

    st.header("6. Export")
    docx_bytes = build_docx(st.session_state["project_fingerprint"], st.session_state["report"], st.session_state["verified_facts"], st.session_state["source_checks"], st.session_state["research_sources"])
    json_bytes = json.dumps({"project": st.session_state["project_fingerprint"], "report": st.session_state["report"], "verified_facts": st.session_state["verified_facts"], "sources": st.session_state["research_sources"]}, indent=2).encode("utf-8")

    e1, e2 = st.columns(2)
    with e1:
        st.download_button("📄 Download Word report", data=docx_bytes, file_name="AHJ_Research_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with e2:
        st.download_button("💾 Save research session", data=json_bytes, file_name="AHJ_Research_Session.json", mime="application/json", use_container_width=True)

st.divider()
st.caption("Research aid only. Code editions, permit requirements, jurisdiction, and AHJ procedures must be confirmed with the applicable authority before construction.")
