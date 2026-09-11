import os
import re
import json
from datetime import datetime, date
from io import BytesIO
from urllib.parse import urlparse

import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt
import urllib.request
import urllib.error

# ============================================================
# AHJ RESEARCH ASSISTANT (Native Search Grounding Version)
# ============================================================

st.set_page_config(
    page_title="AHJ Research Assistant",
    page_icon="🏛️",
    layout="wide",
)

# -----------------------------
# Constants
# -----------------------------
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
🔴 UNKNOWN: Available evidence does not establish theAnswer.
⚠️ VERIFY: A direct AHJ, parcel record, permit record, or code official check is required.
"""

# -----------------------------
# Session State
# -----------------------------
defaults = {
    "report": "",
    "research_log": [],
    "research_sources": [],
    "verified_facts": [],
    "source_checks": [],
    "project_fingerprint": {},
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
        if value:
            return value
    except Exception:
        pass
    return os.getenv(name, default)

GEMINI_KEY = secret_or_env("GEMINI_KEY", "")
# Using the new, supported model
MODEL_NAME = secret_or_env("GEMINI_MODEL", "gemini-3.6-flash") 

# -----------------------------
# Helpers
# -----------------------------
def extract_urls(text):
    if not text:
        return []
    urls = re.findall(r"https?://[^\s\]\)>\"']+", text)
    return list(dict.fromkeys(url.rstrip(".,;:") for url in urls))

def check_url(url, timeout=6):
    result = {"url": url, "reachable": False, "status": "", "note": ""}
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "AHJ-Research-Assistant/3.0"})
        with urllib.request.urlopen(req, timeout=timeout) as response:
            result["reachable"] = 200 <= response.status < 400
            result["status"] = str(response.status)
            result["note"] = "URL responded."
    except urllib.error.HTTPError as exc:
        result["status"] = str(exc.code)
        result["note"] = "Server responded, but access was not successful."
    except Exception as exc:
        result["note"] = type(exc).__name__
    return result

def build_fingerprint(state, project_date, project_type, classification, address, details, scope_answers):
    fp = {
        "state": state,
        "project_date": project_date.isoformat() if isinstance(project_date, date) else str(project_date),
        "project_type": project_type,
        "building_classification": classification,
        "address": address.strip(),
        "details": details.strip(),
    }
    fp.update(scope_answers)
    return fp

def compact_fingerprint(fp):
    lines = []
    for key, value in fp.items():
        if value not in ("", None, "No answer", []):
            label = key.replace("_", " ").title()
            lines.append(f"- {label}: {value}")
    return "\n".join(lines)

def base_system_rules():
    return """
You are an AHJ (Authority Having Jurisdiction) research assistant for construction volunteers.

NON-NEGOTIABLE RESEARCH RULES:
1. LIVE SEARCH REQUIRED: You have access to a live Google Search tool. You MUST use it to search for current code editions, effective dates, AHJ procedures, and official government sources. Do not rely on your internal training memory for time-sensitive information.
2. OFFICIAL SOURCES FIRST: Prefer State building-code agencies, County/City/AHJ portals, Official fire authorities, and Official permit portals.
3. DO NOT HARD-CODE CODE EDITIONS: Never assume the ICC edition is the state's adopted edition. Find the actual adopted code for the specific project date.
4. PROJECT DATE MATTERS: Code applicability depends on permit/application dates, effective dates, and transition provisions.
5. NO INVENTED CODE SECTIONS: Never invent section numbers, permit thresholds, or URLs. If you cannot verify it via search, state UNKNOWN or VERIFY.
6. SOURCE CITATIONS: Every material claim must cite a source. 
7. BE CONCISE: A volunteer should understand the preliminary result in under two minutes.
"""

# -----------------------------
# NATIVE GEMINI CALL WITH SEARCH GROUNDING
# -----------------------------
def call_gemini(prompt, temperature=0.1):
    if not GEMINI_KEY:
        return {"text": "ERROR: GEMINI_KEY is not configured.", "sources": [], "model": "", "error": True}

    try:
        client = genai.Client(api_key=GEMINI_KEY)
        
        # Enable native Google Search Grounding
        config = types.GenerateContentConfig(
            tools=[types.Tool(google_search=types.GoogleSearch())],
            max_output_tokens=4000,
        )
        
        if not MODEL_NAME.startswith(("gemini-1.5", "gemini-2.0", "gemini-3.6")):
            config.temperature = temperature

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
            pass  # Fallback if metadata structure varies

        return {
            "text": text if text else "ERROR: No text output generated.",
            "sources": sources,
            "model": MODEL_NAME,
            "error": False,
        }

    except Exception as exc:
        return {
            "text": f"ERROR: Gemini request failed: {type(exc).__name__}: {exc}",
            "sources": [],
            "model": MODEL_NAME,
            "error": True,
        }

# -----------------------------
# Research Prompts (Updated for Native Search)
# -----------------------------
def discovery_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

RESEARCH PASS 1 — JURISDICTION AND OFFICIAL SOURCE MAP
Use your live Google Search tool to find the authoritative government sources needed to research this project.
Determine:
1. State-level building/code authority.
2. Most likely local permitting authority.
3. Building department / permit portal.
4. Fire/life-safety authority when relevant.
5. Planning/zoning authority when relevant.

Return concise findings with official URLs.
"""

def code_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

RESEARCH PASS 2 — CURRENT APPLICABLE CODE CYCLES
Use your live Google Search tool to search the official state code/building agency first.
Determine for this project and project date:
- building/structural, mechanical, electrical, plumbing, energy, and fire/life-safety codes.
- state/local amendments, effective dates, and mandatory/adoption dates.

For every code family, report: Code family | Adopted edition | Effective/mandatory date | Official source URL.
CRITICAL: Do NOT use model memory. If the official source cannot establish the edition, say UNKNOWN/VERIFY.
"""

def scope_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

RESEARCH PASS 3 — PROJECT-SPECIFIC PERMITS AND CODE TRIGGERS
Use your live Google Search tool to research only the requirements actually relevant to this scope.
For each potentially relevant issue determine:
Issue | Preliminary result | Trigger/fact | Authority | Source URL

Potential categories: building permit, mechanical permit, electrical permit, fire permit, planning/zoning, accessibility, equipment replacement, energy efficiency.
If a permit cannot be established from available evidence, say UNKNOWN/VERIFY and formulate the exact AHJ question.
"""

def conflict_prompt(fp, prior_notes):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

RESEARCH PASS 4 — ERROR CHECK / CONFLICT CHECK
Below are preliminary research notes from other passes:
{prior_notes}

Use your live Google Search tool to look specifically for:
- stale code editions, wrong effective dates, wrong permitting authority, unsupported permit claims.
Return only: 1. Conflict or possible error, 2. Correct/current evidence, 3. Source URL, 4. What the volunteer should do if unresolved.
"""

def synthesis_prompt(fp, research_passes, all_sources):
    source_lines = []
    for i, source in enumerate(all_sources, 1):
        source_lines.append(f"[S{i}] {source.get('title', 'Source')} — {source.get('url', '')}")
    sources_text = "\n".join(source_lines)

    joined = "\n\n".join(f"===== RESEARCH PASS {i + 1} =====\n{item['text']}" for i, item in enumerate(research_passes))

    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

RESEARCH NOTES:
{joined}

SOURCES ACTUALLY RETRIEVED:
{sources_text}

Synthesize the final volunteer-facing report.
- Prefer the strongest authoritative source.
- If notes conflict, resolve them using the source hierarchy and current date.
- Every important code/permit conclusion should cite a source ID such as [S3].
- Use ONLY source IDs that appear in SOURCES ACTUALLY RETRIEVED.

FORMAT EXACTLY:
# PRELIMINARY ANSWER
Start with the practical answer.
| Issue | Result | Confidence |
|---|---|---|

# CURRENT CODE PATH
Give the applicable code families and editions discovered. Cite source IDs.

# WHAT WE KNOW
Maximum 8 bullets. Cite important claims.

# WHAT COULD CHANGE THE ANSWER
Maximum 5 bullets.

# QUESTIONS FOR THE AHJ
Maximum 5 precise questions.

# VOLUNTEER ACTION LIST
Maximum 6 steps, in order.

# SOURCES
List only sources actually retrieved. Use: - [S1] Title — URL

# RESEARCH LIMITATIONS
One short paragraph explaining what could not be independently established.
"""

def clean_report(text):
    if not text:
        return text
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def source_diagnostics(report, sources):
    urls = []
    for source in sources:
        url = source.get("url")
        if url and url not in urls:
            urls.append(url)
    for url in extract_urls(report):
        if url not in urls:
            urls.append(url)
    
    checks = []
    for url in urls[:30]:
        checks.append(check_url(url))
    return checks

def build_docx(fp, report, verified_facts, diagnostics, sources):
    doc = Document()
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)

    doc.add_heading("AHJ Research Report", 0)
    doc.add_paragraph(f"Project: {fp.get('project_type', '')}\nState: {fp.get('state', '')}\nAddress: {fp.get('address', '')}\nProject date: {fp.get('project_date', '')}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
    doc.add_paragraph("DRAFT RESEARCH AID — Verify applicable requirements with the AHJ before relying on this report.")

    doc.add_heading("Project Scope", level=1)
    doc.add_paragraph(compact_fingerprint(fp))

    doc.add_heading("Research Report", level=1)
    for line in report.splitlines():
        line = line.strip()
        if not line: continue
        if line.startswith("# "): doc.add_heading(line[2:], level=1)
        elif line.startswith("## "): doc.add_heading(line[3:], level=2)
        elif line.startswith("### "): doc.add_heading(line[4:], level=3)
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
            status = "Reachable" if item["reachable"] else "Not verified as reachable"
            doc.add_paragraph(f"{status}: {item['url']} ({item.get('status', '')})")

    doc.add_heading("Confidence Guide", level=1)
    doc.add_paragraph(CONFIDENCE_GUIDE)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

def save_session(fp, report, verified_facts, sources):
    return json.dumps({
        "saved": datetime.now().isoformat(),
        "project": fp,
        "report": report,
        "verified_facts": verified_facts,
        "sources": sources,
    }, indent=2)

# ============================================================
# UI
# ============================================================
st.title("🏛️ AHJ Research Assistant")
st.caption("Search-grounded research. Official sources first. State-agnostic code discovery. Human verification for final decisions.")

with st.sidebar:
    st.header("Research Settings")
    depth = st.radio("Research mode", ["Quick Answer", "Deep Research"], index=0)
    st.divider()
    st.subheader("Confidence Guide")
    st.markdown(CONFIDENCE_GUIDE)
    st.divider()
    st.info("This app uses Gemini's native Google Search tool to find live, current code editions and AHJ requirements. No hardcoded data.")

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
elif project_type in ["Parking / Site Improvements", "New Construction", "Addition"]:
    scope["site_work"] = st.text_area("Site work / civil scope", height=100)
else:
    scope["scope_details"] = st.text_area("Additional scope details", height=100)

scope["existing_land_use"] = existing_permit
fp = build_fingerprint(state, project_date, project_type, classification, address, details, scope)
st.session_state["project_fingerprint"] = fp

with st.expander("See the scope fingerprint the research engine will use"):
    st.code(compact_fingerprint(fp))

st.header("3. Research")
research_col1, research_col2 = st.columns([1, 3])
with research_col1:
    run_research = st.button("🔎 Research Project", type="primary", use_container_width=True)
with research_col2:
    st.info("The research engine uses Gemini's native live web search to find current, official sources. Deep Research performs separate jurisdiction, code, scope, and conflict passes.")

if run_research:
    if not address.strip():
        st.error("Enter a project address.")
    elif not GEMINI_KEY:
        st.error("GEMINI_KEY is missing. Add it to Streamlit secrets or your environment variables.")
    else:
        with st.spinner("Searching live web for official sources, code editions, and AHJ requirements..."):
            passes = []

            if depth == "Quick Answer":
                prompt = f"""
{base_system_rules()}
PROJECT: {compact_fingerprint(fp)}
Perform a concise but current research pass using your live Google Search tool.
Identify the authoritative state code source, likely local AHJ, current applicable code path, and key permit/code issues for this scope.
Return the report in the exact format specified in your rules.
"""
                result = call_gemini(prompt)
                passes = [result]
            else:
                prompts = [discovery_prompt(fp), code_prompt(fp), scope_prompt(fp)]
                progress = st.progress(0)

                for i, prompt in enumerate(prompts):
                    result = call_gemini(prompt)
                    passes.append(result)
                    progress.progress((i + 1) / (len(prompts) + 1))

                prior_notes = "\n\n".join(f"PASS {i + 1}:\n{item['text']}" for i, item in enumerate(passes))
                conflict = call_gemini(conflict_prompt(fp, prior_notes))
                passes.append(conflict)
                progress.progress(4 / 5)

                # Aggregate all sources found across all passes
                all_sources = []
                for item in passes:
                    for source in item.get("sources", []):
                        if source not in all_sources:
                            all_sources.append(source)

                synthesis = call_gemini(synthesis_prompt(fp, passes, all_sources))
                passes.append(synthesis)
                progress.progress(1.0)
                result = synthesis

        st.session_state["research_log"] = passes
        
        # Final source aggregation
        final_sources = []
        for item in passes:
            for source in item.get("sources", []):
                if source not in final_sources:
                    final_sources.append(source)
        
        st.session_state["research_sources"] = final_sources
        st.session_state["report"] = clean_report(result["text"])

        with st.spinner("Checking source URLs for basic reachability..."):
            st.session_state["source_checks"] = source_diagnostics(st.session_state["report"], st.session_state["research_sources"])

        st.success(f"Research complete. {len(st.session_state['research_sources'])} live source(s) were retrieved.")

# -----------------------------
# Results
# -----------------------------
if st.session_state["report"]:
    st.header("4. Preliminary Result")
    st.markdown(st.session_state["report"])

    if st.session_state["research_sources"]:
        with st.expander("Sources actually retrieved by Gemini"):
            for i, source in enumerate(st.session_state["research_sources"], 1):
                st.markdown(f"**[S{i}] {source.get('title', 'Source')}**  \n{source.get('url', '')}")

    if st.session_state["source_checks"]:
        with st.expander("Source URL diagnostics"):
            st.caption("These checks only test whether a URL responds. They do not prove that the page supports the AI's claim.")
            for item in st.session_state["source_checks"]:
                if item["reachable"]:
                    st.success(f"Reachable: {item['url']} ({item.get('status', '')})")
                else:
                    st.warning(f"Not independently verified as reachable: {item['url']}")

    st.header("5. Volunteer Verification")
    st.caption("Record facts personally confirmed with an AHJ or official record. These facts can be included in the Word report.")

    with st.form("verification_form"):
        verified_fact = st.text_input("Verified fact")
        verified_source = st.text_input("Source / person / record")
        add_fact = st.form_submit_button("Add verified fact")

        if add_fact and verified_fact.strip():
            st.session_state["verified_facts"].append({
                "fact": verified_fact.strip(),
                "source": verified_source.strip(),
                "date": datetime.now().isoformat(),
            })
            st.success("Added to verification log.")

    if st.session_state["verified_facts"]:
        for i, fact in enumerate(st.session_state["verified_facts"], 1):
            st.markdown(f"**{i}. 🟢 VERIFIED** — {fact['fact']}  \nSource: {fact['source']}")

    st.header("6. Export")
    docx_bytes = build_docx(
        st.session_state["project_fingerprint"],
        st.session_state["report"],
        st.session_state["verified_facts"],
        st.session_state["source_checks"],
        st.session_state["research_sources"],
    )
    json_bytes = save_session(
        st.session_state["project_fingerprint"],
        st.session_state["report"],
        st.session_state["verified_facts"],
        st.session_state["research_sources"],
    ).encode("utf-8")

    e1, e2 = st.columns(2)
    with e1:
        st.download_button("📄 Download Word report", data=docx_bytes, file_name="AHJ_Research_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with e2:
        st.download_button("💾 Save research session", data=json_bytes, file_name="AHJ_Research_Session.json", mime="application/json", use_container_width=True)

st.divider()
st.caption("Research aid only. Code editions, permit requirements, jurisdiction, zoning/CUP conditions, and AHJ procedures must be confirmed with the applicable authority before construction.")
