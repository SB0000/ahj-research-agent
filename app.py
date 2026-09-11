import os
import re
import json
import urllib.request
import urllib.error
from datetime import datetime
from io import BytesIO

import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt

# ============================================================
# AHJ RESEARCH ASSISTANT
# Research-first, concise-output version
# ============================================================

st.set_page_config(
    page_title="AHJ Research Assistant",
    page_icon="🏛️",
    layout="wide",
)

# -----------------------------
# Official reference sources
# -----------------------------
OFFICIAL_SOURCES = {
    "Oregon BCD — Adopted Codes": "https://www.oregon.gov/bcd/codes-stand/pages/adopted-codes.aspx",
    "Oregon BCD — Mechanical": "https://www.oregon.gov/bcd/codes-stand/pages/mechanical.aspx",
    "Oregon BCD — Electrical": "https://www.oregon.gov/bcd/codes-stand/pages/electrical.aspx",
    "Oregon BCD — Commercial Energy": "https://www.oregon.gov/bcd/codes-stand/Pages/energy-commercial-compliance.aspx",
    "Oregon BCD — ePermitting": "https://www.oregon.gov/bcd/epermitting/",
    "Washington County GIS": "https://www.washingtoncountyor.gov/gis",
    "Washington County": "https://www.washingtoncountyor.gov/",
    "City of Hillsboro Building": "https://www.hillsboro-oregon.gov/services/building",
    "Tualatin Valley Fire & Rescue": "https://www.tvfr.com/",
}

# Current Oregon commercial code reference data.
# This is deliberately STATIC and conservative. It is reference data,
# not a substitute for checking the official BCD page.
OREGON_COMMERCIAL_CODES = {
    "OSSC": {
        "name": "Oregon Structural Specialty Code",
        "edition": "2025 OSSC",
        "mandatory": "April 1, 2026",
        "url": OFFICIAL_SOURCES["Oregon BCD — Adopted Codes"],
    },
    "OMSC": {
        "name": "Oregon Mechanical Specialty Code",
        "edition": "2025 OMSC",
        "mandatory": "April 1, 2026",
        "url": OFFICIAL_SOURCES["Oregon BCD — Mechanical"],
    },
    "OEESC": {
        "name": "Oregon Energy Efficiency Specialty Code",
        "edition": "2025 OEESC",
        "mandatory": "July 1, 2025",
        "url": OFFICIAL_SOURCES["Oregon BCD — Commercial Energy"],
    },
    "OESC": {
        "name": "Oregon Electrical Specialty Code",
        "edition": "2023 OESC",
        "mandatory": "Current state electrical cycle; verify effective date with BCD/AHJ",
        "url": OFFICIAL_SOURCES["Oregon BCD — Electrical"],
    },
}

PROJECT_TYPES = [
    "HVAC Replacement",
    "Reroof",
    "Parking / Site Improvements",
    "Interior Remodel",
    "Major Remodel",
    "Addition",
    "New Construction",
    "Other",
]

BUILDING_CLASSIFICATIONS = [
    "Commercial",
    "Assembly",
    "Institutional",
    "Industrial",
    "Agricultural",
    "Residential",
    "Mixed-use",
    "Unknown",
]

CONFIDENCE_GUIDE = """
🟢 VERIFIED
A specific authoritative source was checked and supports the statement.

🟡 LIKELY
Strong preliminary conclusion, but the exact project/AHJ applicability still needs confirmation.

🟠 CONDITIONAL
True only if a stated condition applies.

🔴 UNKNOWN
The available evidence does not establish the answer.

⚠️ VERIFY
A direct AHJ, parcel record, permit record, CUP, or project document check is required.

Important: a source being official does NOT automatically mean every claim made about that source is verified.
"""

# -----------------------------
# Session state
# -----------------------------
defaults = {
    "report": "",
    "research_log": [],
    "verified_facts": [],
    "source_checks": [],
    "project_fingerprint": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# -----------------------------
# Helpers
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
CONFIGURED_MODEL = secret_or_env("GEMINI_MODEL", "gemini-2.5-flash")

MODEL_CANDIDATES = []
for candidate in [
    CONFIGURED_MODEL,
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
]:
    if candidate and candidate not in MODEL_CANDIDATES:
        MODEL_CANDIDATES.append(candidate)


def official_domain(url):
    if not url:
        return False
    u = url.lower()
    return any(
        domain in u
        for domain in [
            "oregon.gov",
            "washingtoncountyor.gov",
            "hillsboro-oregon.gov",
            "tvfr.com",
        ]
    )


def extract_urls(text):
    if not text:
        return []
    urls = re.findall(r"https?://[^\s\]\)>\"']+", text)
    cleaned = []
    for u in urls:
        u = u.rstrip(".,;:")
        if u not in cleaned:
            cleaned.append(u)
    return cleaned


def check_url(url, timeout=6):
    """Checks whether a source URL is reachable.
    This does NOT prove that the page supports a claim.
    """
    result = {
        "url": url,
        "reachable": False,
        "status": "",
        "note": "",
    }
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "AHJ-Research-Assistant/1.0"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as response:
            result["reachable"] = 200 <= response.status < 400
            result["status"] = str(response.status)
            result["note"] = "URL responded."
    except urllib.error.HTTPError as e:
        result["status"] = str(e.code)
        result["note"] = "Server responded, but access was not successful."
    except Exception as e:
        result["note"] = type(e).__name__
    return result


def build_fingerprint(project_type, classification, address, details, scope_answers):
    fp = {
        "project_type": project_type,
        "building_classification": classification,
        "address": address.strip(),
        "details": details.strip(),
    }

    for key, value in scope_answers.items():
        fp[key] = value

    return fp


def compact_fingerprint(fp):
    lines = []
    for key, value in fp.items():
        if value not in ("", None, "No answer", []):
            label = key.replace("_", " ").title()
            lines.append(f"- {label}: {value}")
    return "\n".join(lines)


def call_gemini(prompt, temperature=0.1):
    if not GEMINI_KEY:
        return (
            "ERROR: GEMINI_KEY is not configured. Add GEMINI_KEY to "
            "Streamlit secrets or environment variables."
        )

    genai.configure(api_key=GEMINI_KEY)

    last_error = None

    for model_name in MODEL_CANDIDATES:
        try:
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(
                prompt,
                generation_config={
                    "temperature": temperature,
                },
            )
            text = getattr(response, "text", None)
            if text:
                return text.strip()
        except Exception as exc:
            last_error = exc

    return f"ERROR: Gemini request failed. Last error: {last_error}"


def base_system_rules():
    return f"""
You are an AHJ research assistant for construction volunteers.

Your job is NOT to produce a generic construction-code essay.
Your job is to identify the few requirements that actually matter to the
specific project and point the volunteer to authoritative sources.

RESEARCH PRIORITY:
1. Official state agency source
2. Official county/city/AHJ source
3. Official permit portal or ordinance
4. Official interpretation, adopted code document, CUP/land-use record
5. Other authoritative source
6. General web information
7. Model knowledge only when no better evidence is available

OREGON CODE RULE:
For Oregon commercial projects, check the Oregon BCD adopted-codes source
FIRST. Do not guess a code edition from memory.
Official adopted-code source:
{OFFICIAL_SOURCES["Oregon BCD — Adopted Codes"]}

COMMERCIAL/RESIDENTIAL RULE:
Do not mix residential and commercial code paths.
The user's building classification controls the research path.
If classification is Commercial or Assembly, do NOT present ORSC as the
primary code unless there is a specific reason to believe it applies.

CONFIDENCE:
🟢 VERIFIED = a specific authoritative source was actually checked and supports
the claim.
🟡 LIKELY = strong preliminary conclusion but not fully project/AHJ verified.
🟠 CONDITIONAL = only applies if the stated condition exists.
🔴 UNKNOWN = evidence does not establish it.
⚠️ VERIFY = direct AHJ/property/project confirmation is required.

Never call something VERIFIED merely because you named an official agency.
Do not fabricate quotations, code sections, URLs, permit procedures, or
jurisdiction determinations.

SCOPE DISCIPLINE:
Do not research issues that the project scope explicitly eliminates.
For example, if there are no roof penetrations, do not spend report space
researching roof penetrations. Mark them Not Applicable if useful.

DO NOT PREDICT AN AHJ'S ANSWER.
Instead use:
- What we know
- What remains uncertain
- Exact question to ask the AHJ

Be concise. A volunteer should be able to understand the preliminary answer
in under two minutes.
"""


def quick_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

TASK:
Produce a concise preliminary research answer.

First classify the project and identify which research branches are actually
relevant.

Then give exactly these sections:

# PRELIMINARY ANSWER

A short table with:
Issue | Preliminary result | Confidence

Only include issues relevant to this scope.

# WHAT WE KNOW

Maximum 8 bullets.
Use official sources where possible.
For Oregon codes, use the BCD adopted-codes source first.

# WHAT COULD CHANGE THE ANSWER

Maximum 5 bullets.
Only include facts that would actually change the permit/code/planning result.

# QUESTIONS FOR THE AHJ

Maximum 5 precise questions.
Do not include "expected answers."

# SOURCES

List the most useful sources, with URL and what each source supports.

Do not pad the report with generic fire, plumbing, zoning, structural, parking,
stormwater, or site-work discussion unless the scope makes that relevant.
"""


def deep_pass_prompts(fp):
    project = compact_fingerprint(fp)

    return [
        f"""
{base_system_rules()}

PROJECT:
{project}

RESEARCH PASS 1 — JURISDICTION
Determine the most likely permitting jurisdiction and the exact official
source that should be used to verify it.
Do not call the jurisdiction confirmed unless the source actually establishes
it. Distinguish address context from parcel-level evidence.
Return concise findings and URLs.
""",
        f"""
{base_system_rules()}

PROJECT:
{project}

RESEARCH PASS 2 — CURRENT CODES
This is the highest-priority pass.
Start with the official Oregon BCD adopted-codes page:
{OFFICIAL_SOURCES["Oregon BCD — Adopted Codes"]}

Determine the current applicable commercial code editions for this project.
Pay particular attention to mechanical, electrical, structural/building, and
commercial energy codes.

Do NOT use old editions merely because they appear in model memory.
Give edition, mandatory/effective date when available, official URL, and a
short evidence explanation.
Do not include residential codes unless the project classification requires
them.
""",
        f"""
{base_system_rules()}

PROJECT:
{project}

RESEARCH PASS 3 — PERMITS
Determine which permit categories are genuinely relevant to this exact scope.

For each:
- likely status
- trigger
- what fact would change the answer
- official source
- exact AHJ question if unresolved

Do not assume a permit is required merely because work is commercial.
Do not assume a separate electrical permit is required unless the scope or
official procedure supports that conclusion.
""",
        f"""
{base_system_rules()}

PROJECT:
{project}

RESEARCH PASS 4 — LAND USE / CUP
Only research planning, zoning, setbacks, screening, noise, CUP review, or
site issues that could actually affect this scope.

If an existing CUP is mentioned, identify what needs to be obtained/reviewed.
Do not invent CUP conditions.
Distinguish an existing condition from a guessed local requirement.
""",
        f"""
{base_system_rules()}

PROJECT:
{project}

RESEARCH PASS 5 — SCOPE-SPECIFIC TECHNICAL ISSUES
Research only issues directly created by this scope.

For HVAC replacement, consider only if applicable:
- equipment replacement rules
- electrical MCA/MOP
- disconnect/reconnection
- refrigerant/A2L
- equipment anchorage
- efficiency/energy code
- equipment dimensions/clearances
- condensate or gas connections

Explicitly eliminate issues made irrelevant by the stated scope.
""",
        f"""
{base_system_rules()}

PROJECT:
{project}

RESEARCH PASS 6 — CONFLICTS AND PITFALLS
Look for contradictions, stale code editions, wrong AHJs, residential/commercial
mix-ups, unsupported permit claims, or other statements that a normal Google
search might get wrong.

Return:
- potential conflict
- why it matters
- what authoritative source should settle it
""",
    ]


def synthesis_prompt(fp, research_passes):
    joined = "\n\n".join(
        f"===== RESEARCH PASS {i+1} =====\n{p}"
        for i, p in enumerate(research_passes)
    )

    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

Below are research-pass results. They may contain mistakes. Treat them as
research notes, not truth.

{joined}

Now synthesize the final volunteer-facing report.

IMPORTANT:
- Resolve conflicts in favor of the strongest authoritative source.
- If a source was not actually checked, do not label its claim VERIFIED.
- If the evidence is insufficient, say UNKNOWN or VERIFY.
- Never use "confirmed" just because multiple AI passes agree.
- Do not mix residential and commercial code paths.
- Do not include eliminated scope issues as unresolved issues.
- Do not invent parcel zoning, CUP conditions, fire districts, permit procedures,
  code sections, or effective dates.

FORMAT:

# PRELIMINARY ANSWER

Give the answer first.
Use a compact table:

| Issue | Result | Confidence |
|---|---|---|

# WHAT WE KNOW

Maximum 8 bullets.
Each important code/permit finding should include its source URL.

# WHAT COULD CHANGE THE ANSWER

Maximum 5 bullets.

# QUESTIONS FOR THE AHJ

Maximum 5 questions.

# VOLUNTEER ACTION LIST

Maximum 6 steps, in order.

# SOURCES

Only list sources actually used or specifically worth checking.
For each:
- Source
- URL
- What it supports
- Whether it is authoritative

# RESEARCH LIMITATIONS

Short paragraph explaining anything that could not be independently verified.

Keep the final report compact. Do not repeat the same warning in every section.
"""


def clean_report(text):
    if not text:
        return text

    # Remove accidental duplicated headings.
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Prevent the most dangerous wording pattern.
    text = text.replace("🟢 CONFIRMED", "🟢 VERIFIED")
    return text.strip()


def source_diagnostics(report):
    urls = extract_urls(report)
    checks = []

    for url in urls[:20]:
        checks.append(check_url(url))

    return checks


def build_docx(fp, report, verified_facts, diagnostics):
    doc = Document()

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)

    title = doc.add_heading("AHJ Research Report", 0)
    doc.add_paragraph(
        f"Project: {fp.get('project_type', '')}\n"
        f"Address: {fp.get('address', '')}\n"
        f"Generated: {datetime.now().strftime('%B %d, %Y')}"
    )

    doc.add_paragraph(
        "DRAFT RESEARCH AID — Verify applicable requirements with the AHJ "
        "before relying on this report."
    )

    doc.add_heading("Project Scope", level=1)
    doc.add_paragraph(compact_fingerprint(fp))

    doc.add_heading("Research Report", level=1)

    # Basic markdown-to-docx handling.
    for line in report.splitlines():
        line = line.strip()

        if not line:
            continue

        if line.startswith("# "):
            doc.add_heading(line[2:], level=1)
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("- "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif line.startswith("|"):
            # Keep markdown tables readable in Word rather than trying to
            # perfectly reconstruct every table.
            doc.add_paragraph(line)
        else:
            doc.add_paragraph(line)

    if verified_facts:
        doc.add_heading("Volunteer-Verified Facts", level=1)
        for fact in verified_facts:
            doc.add_paragraph(
                f"{fact.get('fact', '')}\n"
                f"Source: {fact.get('source', '')}",
                style="List Bullet",
            )

    if diagnostics:
        doc.add_heading("Source URL Diagnostics", level=1)
        for item in diagnostics:
            status = "Reachable" if item["reachable"] else "Not verified as reachable"
            doc.add_paragraph(
                f"{status}: {item['url']} ({item.get('status', '')})"
            )

    doc.add_heading("Confidence Guide", level=1)
    doc.add_paragraph(CONFIDENCE_GUIDE)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def save_session(fp, report, notes):
    return json.dumps(
        {
            "saved": datetime.now().isoformat(),
            "project": fp,
            "report": report,
            "verified_facts": notes,
        },
        indent=2,
    )


# ============================================================
# UI
# ============================================================

st.title("🏛️ AHJ Research Assistant")
st.caption(
    "Research first. Official sources first. Concise answers. "
    "Human verification for final permit/code decisions."
)

with st.sidebar:
    st.header("Research Settings")

    depth = st.radio(
        "Research mode",
        ["Quick Answer", "Deep Research"],
        index=0,
    )

    st.divider()

    st.subheader("Confidence Guide")
    st.markdown(CONFIDENCE_GUIDE)

    st.divider()

    st.subheader("Oregon commercial code reference")
    for code_key, code in OREGON_COMMERCIAL_CODES.items():
        st.markdown(
            f"**{code_key}: {code['edition']}**  \n"
            f"Mandatory: {code['mandatory']}"
        )
        st.caption(code["url"])

# -----------------------------
# Project inputs
# -----------------------------
st.header("1. Project")

col1, col2 = st.columns(2)

with col1:
    address = st.text_input(
        "Project address",
        value="24340 NW Meek Rd, 97124",
    )

    project_type = st.selectbox(
        "Project type",
        PROJECT_TYPES,
        index=0,
    )

with col2:
    classification = st.selectbox(
        "Building / use classification",
        BUILDING_CLASSIFICATIONS,
        index=1,
    )

    existing_permit = st.text_input(
        "Existing permit / land-use condition",
        value="Existing Conditional Use Permit (CUP)",
    )

details = st.text_area(
    "Project description",
    value=(
        "Ground-level exterior HVAC unit replacement; like-for-like replacement "
        "with a different brand on an existing exterior pad."
    ),
    height=100,
)

# -----------------------------
# Scope-specific questions
# -----------------------------
st.header("2. Scope Details")

scope = {}

if project_type == "HVAC Replacement":
    a, b = st.columns(2)

    with a:
        scope["equipment_location"] = st.selectbox(
            "Equipment location",
            ["Ground-level exterior", "Rooftop", "Interior", "Other"],
            index=0,
        )

        scope["same_location"] = st.selectbox(
            "Same exact location?",
            ["Yes", "No", "Unknown"],
            index=0,
        )

        scope["existing_pad"] = st.selectbox(
            "Existing pad/slab reused?",
            ["Yes", "No", "Unknown"],
            index=0,
        )

        scope["ductwork"] = st.selectbox(
            "Ductwork affected?",
            ["No", "Yes", "Unknown"],
            index=0,
        )

    with b:
        scope["roof_penetrations"] = st.selectbox(
            "Roof penetrations?",
            ["No", "Yes", "Unknown"],
            index=0,
        )

        scope["site_disturbance"] = st.selectbox(
            "New site disturbance?",
            ["No", "Yes", "Unknown"],
            index=0,
        )

        scope["electrical_changes"] = st.selectbox(
            "Electrical changes?",
            [
                "Unknown",
                "No — same circuit/disconnect",
                "Yes — breaker/wiring/disconnect changes",
            ],
            index=0,
        )

        scope["refrigerant"] = st.selectbox(
            "Replacement refrigerant known?",
            ["Unknown", "A2L / R-32 / R-454B", "Non-A2L", "Other"],
            index=0,
        )

elif project_type == "Reroof":
    scope["roof_type"] = st.selectbox(
        "Roof type",
        ["Unknown", "Low-slope commercial", "Steep-slope", "Other"],
    )
    scope["tear_off"] = st.selectbox(
        "Tear-off or recover?",
        ["Unknown", "Tear-off", "Recover / overlay"],
    )
    scope["equipment_affected"] = st.selectbox(
        "Rooftop equipment affected?",
        ["No", "Yes", "Unknown"],
    )

elif project_type in ["Parking / Site Improvements", "New Construction", "Addition"]:
    scope["site_work"] = st.text_area(
        "Site work / civil scope",
        height=100,
    )

else:
    scope["scope_details"] = st.text_area(
        "Additional scope details",
        height=100,
    )

scope["existing_land_use"] = existing_permit

fp = build_fingerprint(
    project_type,
    classification,
    address,
    details,
    scope,
)

st.session_state["project_fingerprint"] = fp

with st.expander("See the scope fingerprint the research engine will use"):
    st.code(compact_fingerprint(fp))

# -----------------------------
# Research button
# -----------------------------
st.header("3. Research")

research_col1, research_col2 = st.columns([1, 3])

with research_col1:
    run_research = st.button(
        "🔎 Research Project",
        type="primary",
        use_container_width=True,
    )

with research_col2:
    st.info(
        "The research engine prioritizes official state/county/city sources, "
        "uses the project classification to choose the code path, and suppresses "
        "issues eliminated by the scope."
    )

if run_research:
    if not address.strip():
        st.error("Enter a project address.")
    elif not GEMINI_KEY:
        st.error(
            "GEMINI_KEY is missing. Add it to Streamlit secrets "
            "or your environment variables."
        )
    else:
        with st.spinner("Researching official-source paths and project-specific issues..."):
            if depth == "Quick Answer":
                raw = call_gemini(quick_prompt(fp))
                st.session_state["research_log"] = [raw]
            else:
                passes = []
                progress = st.progress(0)

                prompts = deep_pass_prompts(fp)
                for i, prompt in enumerate(prompts):
                    result = call_gemini(prompt)
                    passes.append(result)
                    progress.progress((i + 1) / (len(prompts) + 1))

                raw = call_gemini(synthesis_prompt(fp, passes))
                progress.progress(1.0)

                st.session_state["research_log"] = passes

        st.session_state["report"] = clean_report(raw)

        with st.spinner("Checking source URLs for basic reachability..."):
            st.session_state["source_checks"] = source_diagnostics(
                st.session_state["report"]
            )

        st.success("Research complete.")

# -----------------------------
# Results
# -----------------------------
if st.session_state["report"]:
    st.header("4. Preliminary Result")

    st.markdown(st.session_state["report"])

    # -------------------------
    # Source diagnostics
    # -------------------------
    if st.session_state["source_checks"]:
        with st.expander("Source URL diagnostics"):
            st.caption(
                "These checks only test whether a URL responds. They do not "
                "prove that the page supports the AI's claim."
            )

            for item in st.session_state["source_checks"]:
                if item["reachable"]:
                    st.success(
                        f"Reachable: {item['url']} "
                        f"({item.get('status', '')})"
                    )
                else:
                    st.warning(
                        f"Not independently verified as reachable: {item['url']}"
                    )

    # -------------------------
    # Volunteer verification
    # -------------------------
    st.header("5. Volunteer Verification")

    st.caption(
        "Use this section to record facts you personally confirmed with an AHJ "
        "or official record. These facts can be included in the Word report."
    )

    with st.form("verification_form"):
        verified_fact = st.text_input("Verified fact")
        verified_source = st.text_input("Source / person / record")
        add_fact = st.form_submit_button("Add verified fact")

        if add_fact and verified_fact.strip():
            st.session_state["verified_facts"].append(
                {
                    "fact": verified_fact.strip(),
                    "source": verified_source.strip(),
                    "date": datetime.now().isoformat(),
                }
            )
            st.success("Added to verification log.")

    if st.session_state["verified_facts"]:
        for i, fact in enumerate(st.session_state["verified_facts"], 1):
            st.markdown(
                f"**{i}. 🟢 VERIFIED** — {fact['fact']}  \n"
                f"Source: {fact['source']}"
            )

    # -------------------------
    # Exports
    # -------------------------
    st.header("6. Export")

    docx_bytes = build_docx(
        st.session_state["project_fingerprint"],
        st.session_state["report"],
        st.session_state["verified_facts"],
        st.session_state["source_checks"],
    )

    json_bytes = save_session(
        st.session_state["project_fingerprint"],
        st.session_state["report"],
        st.session_state["verified_facts"],
    ).encode("utf-8")

    e1, e2 = st.columns(2)

    with e1:
        st.download_button(
            "📄 Download Word report",
            data=docx_bytes,
            file_name="AHJ_Research_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with e2:
        st.download_button(
            "💾 Save research session",
            data=json_bytes,
            file_name="AHJ_Research_Session.json",
            mime="application/json",
            use_container_width=True,
        )

# -----------------------------
# Footer
# -----------------------------
st.divider()
st.caption(
    "Research aid only. Code editions, permit requirements, jurisdiction, "
    "zoning/CUP conditions, and AHJ procedures must be confirmed with the "
    "applicable authority before construction."
)
