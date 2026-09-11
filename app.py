import os
import re
import json
from datetime import datetime, date
from io import BytesIO

import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt

# ============================================================
# AHJ RESEARCH ASSISTANT (Native Grounding Edition)
# ============================================================

st.set_page_config(page_title="AHJ Research Assistant", page_icon="🏛️", layout="wide")

# -----------------------------
# State & Config Lists
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
🟢 VERIFIED: A specific authoritative source was actually retrieved.
🟡 LIKELY: Strong preliminary conclusion, but needs AHJ confirmation.
🟠 CONDITIONAL: True only if a stated condition applies.
🔴 UNKNOWN: Available evidence does not establish the answer.
⚠️ VERIFY: A direct AHJ or permit record check is required.
"""

# -----------------------------
# Session state
# -----------------------------
defaults = {
    "report": "", "research_log": [], "research_sources": [],
    "verified_facts": [], "project_fingerprint": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# -----------------------------
# Secrets / Model Config
# -----------------------------
def secret_or_env(name, default=""):
    try:
        value = st.secrets.get(name, None)
        if value: return value
    except Exception: pass
    return os.getenv(name, default)

GEMINI_KEY = secret_or_env("GEMINI_KEY", "")
# Use a REAL model name! gemini-2.0-flash is currently the best/fastest for grounding.
MODEL_NAME = secret_or_env("GEMINI_MODEL", "gemini-2.0-flash") 

# -----------------------------
# Helpers
# -----------------------------
def build_fingerprint(state, project_date, project_type, classification, address, details, scope_answers):
    fp = {
        "state": state, "project_date": str(project_date), "project_type": project_type,
        "building_classification": classification, "address": address.strip(), "details": details.strip(),
    }
    for key, value in scope_answers.items(): fp[key] = value
    return fp

def compact_fingerprint(fp):
    lines = [f"- {k.replace('_', ' ').title()}: {v}" for k, v in fp.items() if v not in ("", None, "No answer", [])]
    return "\n".join(lines)

def base_system_rules():
    return """
You are an AHJ research assistant. You have access to LIVE GOOGLE SEARCH. You MUST use it.

NON-NEGOTIABLE RESEARCH RULES:
1. USE YOUR SEARCH TOOL: Do not rely on your training memory for code editions, permit procedures, or zoning. Search the live web for the specific address, state, and project type.
2. OFFICIAL SOURCES FIRST: Prefer .gov, .us, and official city/county portals.
3. NO INVENTED CODE SECTIONS: Never invent section numbers, permit thresholds, or effective dates. If you cannot verify it via search, say UNKNOWN or VERIFY.
4. SOURCE CITATIONS: Cite your sources clearly in the text using URLs.
"""

def call_gemini_with_search(prompt, temperature=0.1):
    if not GEMINI_KEY:
        return {"text": "ERROR: GEMINI_KEY is not configured.", "sources": [], "error": True}

    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config_kwargs = {
            "max_output_tokens": 4000,
            # THIS IS THE MAGIC: Native Google Search Grounding
            "tools": [types.Tool(google_search=types.GoogleSearch())]
        }
        
        if not MODEL_NAME.startswith("gemini-2."):
            config_kwargs["temperature"] = temperature

        response = client.models.generate_content(
            model=MODEL_NAME, contents=prompt,
            config=types.GenerateContentConfig(**config_kwargs),
        )

        text = getattr(response, "text", "")
        sources = []
        
        # Extract native grounding sources from the new SDK metadata
        try:
            if response.candidates and response.candidates[0].grounding_metadata:
                metadata = response.candidates[0].grounding_metadata
                if hasattr(metadata, 'grounding_chunks'):
                    for chunk in metadata.grounding_chunks:
                        if hasattr(chunk, 'web') and chunk.web:
                            sources.append({"title": chunk.web.title, "url": chunk.web.uri})
        except Exception:
            pass # Fallback to text parsing if SDK structure varies
            
        # Fallback: Extract URLs written in the text by the model
        if not sources:
            urls = re.findall(r"https?://[^\s\]\)>\"']+", text)
            for url in set(urls): sources.append({"title": "Cited in text", "url": url.rstrip(".,;:")})

        return {"text": text.strip(), "sources": sources, "model": MODEL_NAME, "error": False}

    except Exception as exc:
        return {
            "text": f"ERROR: Gemini request failed: {type(exc).__name__}: {exc}\n\n*(Note: If you see 'Resource Exhausted', you hit the free-tier limit. Wait 60 seconds and try again.)*",
            "sources": [], "model": MODEL_NAME, "error": True,
        }

# -----------------------------
# Prompts
# -----------------------------
def quick_prompt(fp):
    return f"""{base_system_rules()}
PROJECT: {compact_fingerprint(fp)}
Perform a concise but current research pass using your Google Search tool.
Identify the authoritative state code source, the likely local AHJ, and the current applicable code path.
Format exactly with headers: # PRELIMINARY ANSWER (Table), # CURRENT CODE PATH, # WHAT WE KNOW, # WHAT COULD CHANGE THE ANSWER, # QUESTIONS FOR THE AHJ, # SOURCES.
"""

def deep_prompts(fp):
    base = base_system_rules() + f"\nPROJECT: {compact_fingerprint(fp)}\n"
    p1 = base + "PASS 1: Search for Jurisdiction, Local AHJ, and Official Source Map for this address."
    p2 = base + "PASS 2: Search for the CURRENT APPLICABLE CODE CYCLES (Building, Mechanical, Electrical, Energy) and their effective dates for this project date."
    p3 = base + "PASS 3: Search for PROJECT-SPECIFIC PERMITS AND CODE TRIGGERS based on the scope."
    return [p1, p2, p3]

def synthesis_prompt(fp, passes):
    joined = "\n\n".join(f"===== PASS {i+1} =====\n{p['text']}" for i, p in enumerate(passes))
    return f"""{base_system_rules()}
PROJECT: {compact_fingerprint(fp)}
RESEARCH NOTES: {joined}
Synthesize the final volunteer-facing report. Use the exact headers requested in the quick prompt format.
"""

# -----------------------------
# UI
# -----------------------------
st.title("🏛️ AHJ Research Assistant")
st.caption("Native Google Search Grounding. Official sources first.")

with st.sidebar:
    st.header("Settings")
    depth = st.radio("Research mode", ["Quick Answer", "Deep Research"], index=0)
    st.divider()
    st.subheader("Confidence Guide")
    st.markdown(CONFIDENCE_GUIDE)

st.header("1. Project")
col1, col2 = st.columns(2)
with col1:
    state = st.selectbox("State", STATE_OPTIONS, index=STATE_OPTIONS.index("Oregon"))
    address = st.text_input("Address", value="24340 NW Meek Rd, 97124")
    project_date = st.date_input("Project Date", value=date.today())
    project_type = st.selectbox("Project Type", PROJECT_TYPES, index=0)
with col2:
    classification = st.selectbox("Classification", BUILDING_CLASSIFICATIONS, index=1)
    existing_permit = st.text_input("Existing Permit/CUP", value="Existing CUP")
    building_status = st.selectbox("Building Status", ["Existing building", "New construction", "Addition"], index=0)
    details = st.text_area("Description", value="Ground-level exterior HVAC unit replacement; like-for-like.", height=100)

scope = {"building_status": building_status, "existing_land_use": existing_permit}
fp = build_fingerprint(state, project_date, project_type, classification, address, details, scope)

st.header("2. Research")
if st.button("🔎 Research Project", type="primary", use_container_width=True):
    if not GEMINI_KEY:
        st.error("GEMINI_KEY missing in secrets.")
    else:
        with st.spinner("Searching live web and analyzing..."):
            passes = []
            if depth == "Quick Answer":
                res = call_gemini_with_search(quick_prompt(fp))
                passes = [res]
                final_res = res
            else:
                prompts = deep_prompts(fp)
                prog = st.progress(0)
                for i, p in enumerate(prompts):
                    res = call_gemini_with_search(p)
                    passes.append(res)
                    prog.progress((i+1)/4)
                
                final_res = call_gemini_with_search(synthesis_prompt(fp, passes))
                passes.append(final_res)
                prog.progress(1.0)

            st.session_state["report"] = final_res["text"]
            st.session_state["research_sources"] = final_res["sources"]
            st.session_state["research_log"] = passes

if st.session_state["report"]:
    st.header("3. Result")
    st.markdown(st.session_state["report"])
    
    if st.session_state["research_sources"]:
        with st.expander("Live Sources Retrieved via Google Search"):
            for i, s in enumerate(st.session_state["research_sources"], 1):
                st.markdown(f"**[{i}] {s.get('title', 'Source')}**  \n{s.get('url', '')}")

    # Docx Export
    st.header("4. Export")
    doc = Document()
    doc.add_heading("AHJ Research Report", 0)
    doc.add_paragraph(f"Project: {fp['project_type']}\nAddress: {fp['address']}\nGenerated: {datetime.now().strftime('%B %d, %Y')}")
    for line in st.session_state["report"].splitlines():
        if line.startswith("# "): doc.add_heading(line[2:], level=1)
        elif line.startswith("## "): doc.add_heading(line[3:], level=2)
        else: doc.add_paragraph(line)
        
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    
    st.download_button("📄 Download Word Report", data=buf.getvalue(), file_name="AHJ_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
