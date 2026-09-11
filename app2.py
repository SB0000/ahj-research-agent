````python
"""
AHJ Research Agent
Kingdom Hall Project Research Assistant
Version 2.0

Purpose:
    Help volunteers research AHJ requirements for Kingdom Hall
    construction, renovation, reroofing, HVAC, site work, etc.

IMPORTANT:
    This application is a research assistant, NOT a code official,
    attorney, permit expediter, engineer, architect, or AHJ.

    AI-generated information is never treated as verified merely
    because the AI sounds confident.

Confidence labels:
    🟢 VERIFIED       = supported by a specific authoritative source
                        supplied/cited in the research.
    🟡 AI KNOWLEDGE   = plausible information from model/reference
                        knowledge that has not been independently verified.
    🟠 UNCERTAIN      = information may be outdated, incomplete,
                        conditional, or jurisdiction-dependent.
    🔴 CANNOT VERIFY  = could not establish the answer.
    ⚠️ VERIFY         = volunteer should confirm directly with AHJ.

The strongest part of this application is deliberately NOT the AI's
confidence. It is the separation between:
    1. What we know
    2. What we think
    3. What needs verification
    4. What source supports the finding
    5. What the volunteer should ask the AHJ
"""

import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
import json
import datetime
import time
import re
from urllib.parse import urlparse


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="AHJ Research Agent",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# APPLICATION CONSTANTS
# ============================================================

APP_VERSION = "2.0"

CONFIDENCE_GUIDE = {
    "VERIFIED": "🟢 VERIFIED",
    "AI_KNOWLEDGE": "🟡 AI KNOWLEDGE",
    "UNCERTAIN": "🟠 UNCERTAIN",
    "CANNOT_VERIFY": "🔴 CANNOT VERIFY",
    "VERIFY": "⚠️ VERIFY",
}

# These are intentionally reference points, not a claim that they
# are current forever. The UI tells the user when the reference was
# last reviewed.
#
# Expand this over time for states where your organization frequently
# works.
REFERENCE_LAST_REVIEWED = "September 2026"

STATE_CODE_REFERENCE = {
    "oregon": {
        "state": "Oregon",
        "official_agency": "Oregon Building Codes Division (BCD)",
        "official_domains": [
            "oregon.gov",
        ],
        "codes": [
            {
                "discipline": "Building",
                "name": "Oregon Structural Specialty Code",
                "edition": "2025",
                "effective": "Mandatory April 1, 2026",
                "source_url": "https://www.oregon.gov/bcd/codes-stand/Pages/adopted-codes.aspx",
            },
            {
                "discipline": "Mechanical",
                "name": "Oregon Mechanical Specialty Code",
                "edition": "2025",
                "effective": "Mandatory April 1, 2026",
                "source_url": "https://www.oregon.gov/bcd/codes-stand/pages/mechanical.aspx",
            },
            {
                "discipline": "Energy",
                "name": "Oregon Energy Efficiency Specialty Code",
                "edition": "2025",
                "effective": "Mandatory July 1, 2025",
                "source_url": "https://www.oregon.gov/bcd/codes-stand/Pages/energy-commercial-compliance.aspx",
            },
            {
                "discipline": "Electrical",
                "name": "Oregon Electrical Specialty Code",
                "edition": "2023",
                "effective": "Current reference; verify with BCD/AHJ",
                "source_url": "https://www.oregon.gov/bcd/codes-stand/pages/electrical.aspx",
            },
            {
                "discipline": "Plumbing",
                "name": "Oregon Plumbing Specialty Code",
                "edition": "2023",
                "effective": "Verify current adoption",
                "source_url": "https://www.oregon.gov/bcd/codes-stand/pages/plumbing.aspx",
            },
            {
                "discipline": "Fire",
                "name": "Oregon Fire Code / Fire & Life Safety requirements",
                "edition": "Verify current edition and local fire AHJ",
                "effective": "Verify with fire AHJ",
                "source_url": "https://www.oregon.gov/",
            },
        ],
    },
    "washington": {
        "state": "Washington",
        "official_agency": "Washington State Building Code Council",
        "official_domains": [
            "sbcc.wa.gov",
            "wa.gov",
        ],
        "codes": [
            {
                "discipline": "Building",
                "name": "Washington State Building Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://sbcc.wa.gov/",
            },
            {
                "discipline": "Mechanical",
                "name": "Washington State Mechanical Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://sbcc.wa.gov/",
            },
            {
                "discipline": "Energy",
                "name": "Washington State Energy Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://sbcc.wa.gov/",
            },
            {
                "discipline": "Electrical",
                "name": "Washington Electrical Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://lni.wa.gov/",
            },
        ],
    },
    "california": {
        "state": "California",
        "official_agency": "California Building Standards Commission",
        "official_domains": [
            "dgs.ca.gov",
            "ca.gov",
        ],
        "codes": [
            {
                "discipline": "Building",
                "name": "California Building Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://www.dgs.ca.gov/BSC",
            },
            {
                "discipline": "Mechanical",
                "name": "California Mechanical Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://www.dgs.ca.gov/BSC",
            },
            {
                "discipline": "Energy",
                "name": "California Energy Code",
                "edition": "Verify current effective edition",
                "effective": "Verify",
                "source_url": "https://www.energy.ca.gov/",
            },
        ],
    },
}


# ============================================================
# PROJECT KNOWLEDGE BASE
# ============================================================

PROJECT_TYPES = {
    "HVAC Replacement": {
        "description": "Replacement of existing HVAC equipment",
        "questions": [
            "Is the equipment ground-mounted, interior, or rooftop?",
            "Is the replacement truly like-for-like?",
            "Will equipment location change?",
            "Will weight, dimensions, curb, or supports change?",
            "Will ductwork change?",
            "Will refrigerant type change?",
            "Will electrical circuits, disconnects, breakers, or wiring change?",
            "Will roof penetrations change?",
            "Will equipment screening or enclosures change?",
        ],
        "concerns": [
            "Mechanical permit",
            "Electrical permit if electrical work changes",
            "Energy-code requirements for replacement equipment",
            "Structural review if equipment/support conditions change",
            "Rooftop equipment anchorage",
            "Refrigerant requirements",
            "A2L refrigerant requirements if applicable",
            "Ductwork and fire/life-safety implications",
            "Noise requirements",
            "Equipment screening",
            "Roof penetrations and waterproofing",
        ],
    },
    "Reroof": {
        "description": "Replacement or alteration of existing roof system",
        "questions": [
            "Is this tear-off, recover, or both?",
            "Will the roof covering type change?",
            "Will insulation change?",
            "Will the roof deck be replaced?",
            "Will rooftop equipment be moved?",
            "Will roof penetrations change?",
            "Will cranes, dumpsters, or materials occupy public right-of-way?",
        ],
        "concerns": [
            "Roofing/building permit",
            "Energy-code requirements",
            "Structural review",
            "Roof deck condition",
            "Wind/uplift requirements",
            "Rooftop equipment",
            "Fall protection",
            "Crane/staging requirements",
            "Right-of-way use",
            "Fire/life-safety requirements",
        ],
    },
    "Parking Lot Improvements": {
        "description": "Parking lot modification, expansion, or improvement",
        "questions": [
            "Is the parking footprint increasing, decreasing, or unchanged?",
            "Will impervious area increase?",
            "Will grading change?",
            "Will drainage/stormwater facilities change?",
            "Will driveway access change?",
            "Will parking spaces or accessible spaces change?",
            "Will lighting change?",
            "Will landscaping or screening change?",
            "Will work occur in public right-of-way?",
        ],
        "concerns": [
            "Planning/zoning review",
            "Civil/site review",
            "Stormwater",
            "Grading/drainage",
            "ADA/accessibility",
            "Driveway access",
            "Right-of-way",
            "Lighting",
            "Landscaping",
            "Parking count",
            "Erosion control",
        ],
    },
    "Interior Remodel": {
        "description": "Interior remodel with limited or no structural changes",
        "questions": [
            "Are walls being removed or added?",
            "Are exits or exit paths changing?",
            "Are rooms or occupancy areas changing?",
            "Are plumbing fixtures changing?",
            "Are electrical circuits changing?",
            "Are HVAC zones or equipment changing?",
            "Are fire sprinklers changing?",
            "Is accessibility affected?",
        ],
        "concerns": [
            "Building permit",
            "Occupancy classification",
            "Means of egress",
            "Accessibility",
            "Fire/life safety",
            "Electrical permit",
            "Mechanical permit",
            "Plumbing permit",
            "Fire sprinkler review",
            "Energy requirements",
        ],
    },
    "Major Remodel": {
        "description": "Substantial remodel of an existing building",
        "questions": [
            "What percentage of the building is affected?",
            "Are structural elements changing?",
            "Are exits changing?",
            "Are occupancy/use changes involved?",
            "Is the building envelope changing?",
            "Are mechanical systems changing?",
            "Are electrical systems changing?",
            "Are plumbing systems changing?",
            "Are fire protection systems changing?",
        ],
        "concerns": [
            "Building permit",
            "Existing-building-code provisions",
            "Occupancy classification",
            "Accessibility",
            "Structural review",
            "Fire/life safety",
            "Energy code",
            "Mechanical",
            "Electrical",
            "Plumbing",
            "Fire protection",
        ],
    },
    "Addition": {
        "description": "Addition to an existing building",
        "questions": [
            "What is the proposed addition square footage?",
            "Will foundation work be required?",
            "Will the addition change occupancy?",
            "Will utilities need modification?",
            "Will stormwater change?",
            "Will setbacks or zoning standards be affected?",
            "Will parking be affected?",
            "Will fire access change?",
        ],
        "concerns": [
            "Building permit",
            "Planning/zoning",
            "Setbacks",
            "Lot coverage",
            "Structural",
            "Foundation",
            "Energy code",
            "Accessibility",
            "Fire access",
            "Parking",
            "Stormwater",
            "Utilities",
        ],
    },
    "New Construction": {
        "description": "New Kingdom Hall construction",
        "questions": [
            "Has the parcel been confirmed?",
            "Is the proposed use permitted?",
            "What is the proposed building size?",
            "Are utilities available?",
            "What is the proposed parking count?",
            "Will public improvements be required?",
            "Are wetlands, floodplain, or environmental constraints present?",
            "Are fire access requirements known?",
        ],
        "concerns": [
            "Land use approval",
            "Building permit",
            "Site development",
            "Civil engineering",
            "Stormwater",
            "Utilities",
            "Fire access",
            "Parking",
            "Accessibility",
            "Environmental overlays",
            "Right-of-way",
            "Public improvements",
        ],
    },
}


# ============================================================
# API CONFIGURATION
# ============================================================

def configure_api():
    """Configure Gemini from Streamlit secrets."""
    try:
        api_key = st.secrets.get("GEMINI_KEY", "")

        if not api_key:
            return False

        genai.configure(api_key=api_key)
        return True

    except Exception:
        return False


api_ready = configure_api()


# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def now_iso():
    return datetime.datetime.now().isoformat(timespec="seconds")


def detect_state(address):
    """Best-effort state detection from address text."""
    address_lower = address.lower()

    state_patterns = {
        "oregon": [
            r"\boregon\b",
            r"\bor\b",
        ],
        "washington": [
            r"\bwashington\b",
            r"\bwa\b",
        ],
        "california": [
            r"\bcalifornia\b",
            r"\bca\b",
        ],
    }

    for state, patterns in state_patterns.items():
        for pattern in patterns:
            if re.search(pattern, address_lower):
                return state

    return None


def get_reference_context(address):
    """Return static reference data when available."""
    state = detect_state(address)

    if not state or state not in STATE_CODE_REFERENCE:
        return {
            "state": None,
            "text": (
                "No static state code reference is available for this "
                "address. Do not invent code-cycle information."
            ),
        }

    data = STATE_CODE_REFERENCE[state]

    lines = [
        f"STATE REFERENCE: {data['state']}",
        f"Official agency: {data['official_agency']}",
        f"Reference reviewed: {REFERENCE_LAST_REVIEWED}",
        "",
    ]

    for code in data["codes"]:
        lines.append(
            f"- {code['discipline']}: {code['name']} — "
            f"Edition: {code['edition']} — "
            f"Effective: {code['effective']} — "
            f"Source: {code['source_url']}"
        )

    return {
        "state": state,
        "text": "\n".join(lines),
    }


def extract_urls(text):
    """Extract URLs from AI output."""
    if not text:
        return []

    urls = re.findall(
        r"https?://[^\s<>\]\)\"']+",
        text
    )

    cleaned = []

    for url in urls:
        url = url.rstrip(".,;:")
        if url not in cleaned:
            cleaned.append(url)

    return cleaned


def is_official_domain(url, allowed_domains=None):
    """Basic domain check. This does NOT prove content accuracy."""
    try:
        hostname = urlparse(url).hostname or ""
        hostname = hostname.lower()

        if allowed_domains:
            return any(
                hostname == domain or hostname.endswith("." + domain)
                for domain in allowed_domains
            )

        return (
            hostname.endswith(".gov")
            or hostname.endswith(".gov.au")
            or hostname.endswith(".gov.uk")
        )

    except Exception:
        return False


def clean_ai_text(text):
    """Remove accidental model artifacts."""
    if not text:
        return ""

    text = text.replace("```markdown", "")
    text = text.replace("```", "")

    return text.strip()


def safe_json_load(text):
    """Attempt to parse JSON returned by Gemini."""
    if not text:
        return None

    text = text.strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    match = re.search(r"\{.*\}", text, re.DOTALL)

    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            return None

    return None


# ============================================================
# GEMINI CALL
# ============================================================

def call_ai(prompt, max_retries=2):
    """
    Call Gemini with retry handling.

    NOTE:
        This function does NOT claim that Gemini has live web access.
        If a source is cited by the model, the application treats it
        as a cited source, not automatically as verified truth.
    """

    if not api_ready:
        return (
            "ERROR: Gemini API is not configured. "
            "Add GEMINI_KEY to Streamlit secrets."
        )

    model_name = "gemini-3.6-flash"

    try:
        model = genai.GenerativeModel(model_name)
    except Exception as exc:
        return f"ERROR: Could not initialize Gemini: {exc}"

    for attempt in range(max_retries + 1):
        try:
            response = model.generate_content(prompt)

            if response and getattr(response, "text", None):
                return response.text.strip()

            return "ERROR: Gemini returned no usable response."

        except Exception as exc:
            error_text = str(exc)

            if "429" in error_text or "RESOURCE_EXHAUSTED" in error_text:
                if attempt < max_retries:
                    time.sleep(5 * (attempt + 1))
                    continue

                return (
                    "ERROR: Gemini rate limit reached. "
                    "Please wait and try again."
                )

            if "403" in error_text:
                return (
                    "ERROR: Gemini API authentication failed. "
                    "Check GEMINI_KEY."
                )

            if "404" in error_text:
                return (
                    "ERROR: Gemini model is unavailable. "
                    "Update MODEL_NAME in app.py."
                )

            if attempt < max_retries:
                time.sleep(2)
                continue

            return f"ERROR: Gemini request failed: {error_text}"

    return "ERROR: Unable to contact Gemini."


# ============================================================
# PROMPT BUILDING
# ============================================================

def build_research_prompt(
    address,
    project_type,
    project_details,
    selected_concerns,
):
    """
    Build the master research prompt.

    The prompt deliberately prevents the model from converting
    AI knowledge into fake certainty.
    """

    reference = get_reference_context(address)

    concerns_text = "\n".join(
        f"- {item}" for item in selected_concerns
    )

    project_questions = PROJECT_TYPES.get(
        project_type,
        {}
    ).get("questions", [])

    questions_text = "\n".join(
        f"- {item}" for item in project_questions
    )

    prompt = f"""
You are an AHJ research assistant for volunteers working on
Kingdom Hall construction and renovation projects.

Your job is NOT to give legal advice and NOT to pretend to be
the Authority Having Jurisdiction.

Your job is to create a practical research package that helps
a volunteer determine:

1. Which jurisdictions and AHJs matter.
2. Which codes appear relevant.
3. Which permits may be involved.
4. Which planning/zoning issues should be investigated.
5. Which site/public-works issues should be investigated.
6. Which project-specific issues could create surprises.
7. What remains unknown.
8. Exactly what the volunteer should ask the AHJ.

============================================================
CRITICAL ACCURACY RULE
============================================================

Building codes, permit requirements, zoning rules, agency names,
jurisdiction boundaries, and procedures change.

NEVER convert model knowledge into a claim of verification.

Use these labels:

🟢 VERIFIED
Only when a specific authoritative source is identified and the
source actually appears relevant to the statement.

🟡 AI KNOWLEDGE
Information you believe is probably correct based on general
knowledge, but which you have NOT independently established from
a current authoritative source.

🟠 UNCERTAIN
Information that may be outdated, conditional, incomplete, or
dependent on facts that are not known.

🔴 CANNOT VERIFY
You cannot establish the answer from the information available.

⚠️ VERIFY
The volunteer should directly confirm this with the appropriate AHJ.

IMPORTANT:
- Do not use 🟢 VERIFIED merely because something sounds certain.
- Do not treat your training data as a current code book.
- Do not invent code sections.
- Do not invent permit requirements.
- Do not invent zoning designations.
- Do not invent parcel information.
- Do not invent agency names or phone numbers.
- Do not claim an address is inside a city or county unless you have
  evidence for that conclusion.
- Do not fabricate URLs.
- Do not say "the AHJ will require..." unless supported.
- Do not provide an "Expected Answer" to an AHJ question.
- When uncertain, create an AHJ question instead.

============================================================
STATIC REFERENCE DATA
============================================================

The following reference information is available to you.

IMPORTANT:
This reference file is NOT itself considered a live verification.
Treat information from it as 🟡 AI KNOWLEDGE unless a specific
authoritative source is separately cited and supports the claim.

{reference["text"]}

============================================================
PROJECT
============================================================

Address:
{address}

Project type:
{project_type}

Project details:
{project_details if project_details else "No additional project details supplied."}

Project-specific questions worth considering:
{questions_text}

Known project concerns:
{concerns_text}

============================================================
RESEARCH METHOD
============================================================

Think in terms of evidence.

For every important finding, distinguish:

WHAT WE KNOW
A fact directly supported by a source or supplied project information.

WHAT WE THINK
A reasonable inference, general practice, or model knowledge.

WHAT WE CANNOT ESTABLISH
Information that needs parcel research, current code research,
or direct AHJ confirmation.

WHAT THE VOLUNTEER SHOULD ASK
A precise question that will close the information gap.

============================================================
JURISDICTION
============================================================

Determine, if possible:

- City or unincorporated area
- County
- Building department
- Planning department
- Fire authority
- Public works / engineering authority
- Electrical authority
- Utility providers or districts if relevant
- Other special districts if relevant

Do NOT guess.

If the exact jurisdiction cannot be established from the available
information, explicitly say:

"Jurisdiction requires verification."

============================================================
CODES
============================================================

Identify potentially applicable:

- Building
- Existing building
- Mechanical
- Electrical
- Plumbing
- Energy
- Fire/life safety
- Accessibility
- Local amendments

For each code, provide:

- Discipline
- Code name
- Edition if known
- Effective date if known
- Confidence
- Source URL if known
- What still needs verification

If the edition cannot be established reliably, say so.

============================================================
PERMITS
============================================================

Analyze likely permit categories:

- Building
- Mechanical
- Electrical
- Plumbing
- Fire
- Planning
- Site/civil
- Right-of-way
- Other

For each:

- Status: likely / unlikely / conditional / unknown
- Why
- Triggering condition
- Confidence
- AHJ question

Never say "permit required" solely because that is common practice.

============================================================
PLANNING AND ZONING
============================================================

Determine what should be checked:

- Zoning
- Permitted use
- Conditional use
- Setbacks
- Lot coverage
- Parking
- Landscaping
- Screening
- Signage
- Noise
- Historic resources
- Environmental overlays
- Floodplain
- Wetlands
- Access
- Easements
- Development standards

Do not invent the property's actual zoning designation.

============================================================
SITE AND PUBLIC WORKS
============================================================

Only include issues relevant to this project.

Consider:

- Right-of-way
- Driveway access
- Sidewalk
- Stormwater
- Grading
- Drainage
- Utilities
- Street occupancy
- Crane/staging
- Dumpster placement
- Temporary construction access
- Public improvements

For a small interior-only project, do not overwhelm the volunteer
with irrelevant site-development requirements.

============================================================
PROJECT-SPECIFIC LOGIC
============================================================

Use conditional reasoning.

Examples:

IF rooftop equipment changes:
    investigate structural review, anchorage, curb/supports,
    roof penetrations, equipment dimensions and weight.

IF equipment location changes:
    investigate structural, mechanical, electrical, planning,
    noise and screening implications as applicable.

IF electrical circuit/disconnect/wiring changes:
    investigate electrical permit and electrical scope.

IF refrigerant changes:
    investigate refrigerant-specific requirements.

IF A2L refrigerant may be involved:
    explicitly flag A2L applicability for AHJ/code verification.

IF ductwork changes:
    investigate mechanical and fire/life-safety implications.

IF parking or impervious area changes:
    investigate planning, stormwater, civil, accessibility,
    and right-of-way issues.

Do not claim a conditional issue applies unless the project facts
support the condition.

============================================================
DIG DEEPER
============================================================

Identify unresolved issues that could materially affect:

- Permit requirements
- Project cost
- Schedule
- Design
- Inspections
- Site work
- AHJ coordination

Prioritize them.

============================================================
QUESTIONS FOR AHJ
============================================================

Create precise questions grouped by:

- Building
- Mechanical
- Electrical
- Plumbing
- Planning/Zoning
- Fire
- Public Works/Engineering
- Other

For EACH question include:

1. Exact question to ask
2. Why it matters
3. What information the volunteer needs to learn
4. Follow-up question if appropriate

DO NOT include "Expected Answer."

============================================================
VOLUNTEER NEXT STEPS
============================================================

Create a prioritized checklist.

Use:

1. First
2. Next
3. Then
4. Before permit submission
5. Before construction

============================================================
LIMITED INFORMATION
============================================================

If something cannot be determined, do not fill the gap with a guess.

Say:

"Not established from available information."

Then tell the volunteer exactly how to verify it.

============================================================
OUTPUT FORMAT
============================================================

Return ONLY the following report.

# AHJ RESEARCH REPORT

## 📊 EXECUTIVE SUMMARY

Include:
- Project
- Address
- Overall research status
- Biggest unresolved issues
- Most important AHJ contacts

## 📍 JURISDICTION

## 📚 APPLICABLE CODES

## 📋 PERMIT REQUIREMENTS

## 🏘️ PLANNING AND ZONING

## 🚜 SITE AND PUBLIC WORKS

## 🔍 SCOPE-SPECIFIC CONSIDERATIONS

## ⚠️ DIG DEEPER

## ❓ QUESTIONS FOR AHJ

## ✅ VOLUNTEER NEXT STEPS

## 📞 IF INFORMATION IS LIMITED

## 📚 SOURCES

For each source provide:

- Source name
- URL
- What it supports
- Whether it appears authoritative
- What the volunteer should verify

============================================================
FINAL SAFETY RULE
============================================================

The purpose of this report is to make the volunteer better prepared
before contacting the AHJ.

It is NOT permission to proceed with construction.

Do not substitute AI confidence for AHJ confirmation.
"""

    return prompt


# ============================================================
# RESEARCH
# ============================================================

def quick_research(
    address,
    project_type,
    project_details,
    concerns,
):
    prompt = build_research_prompt(
        address,
        project_type,
        project_details,
        concerns,
    )

    return call_ai(prompt)


def deep_research(
    address,
    project_type,
    project_details,
    concerns,
):
    """
    Multi-pass research.

    These are independent research passes followed by synthesis.
    They are intentionally separated so the final report is not
    produced from one giant unconstrained prompt.
    """

    results = {}

    state_reference = get_reference_context(address)

    base_context = f"""
Address: {address}
Project type: {project_type}
Project details: {project_details or "Not provided"}

Static state reference:
{state_reference["text"]}

IMPORTANT:
The model must not treat model knowledge as verified.
"""

    research_questions = {
        "jurisdiction": f"""
Determine the likely jurisdiction and AHJs for:

{base_context}

Research conceptually:
- city vs unincorporated county
- county
- building department
- planning
- fire
- public works
- electrical
- utilities
- special districts

Do not guess.
Clearly separate facts from uncertainty.
Provide source URLs only when you know them.
""",

        "codes": f"""
Identify potentially applicable current code families for:

{base_context}

Consider:
- building
- existing building
- mechanical
- electrical
- plumbing
- energy
- fire
- accessibility
- local amendments

For each code:
- name
- edition if known
- effective date if known
- confidence
- source
- what must be verified

Do not invent editions or code sections.
""",

        "permits": f"""
Analyze potential permits for:

{base_context}

Consider:
- building
- mechanical
- electrical
- plumbing
- fire
- planning
- site/civil
- right-of-way

Explain the trigger for each permit.
Do not convert typical practice into certainty.
""",

        "planning": f"""
Analyze planning/zoning issues for:

{base_context}

Consider:
- zoning
- permitted use
- setbacks
- parking
- landscaping
- screening
- signage
- noise
- historic
- floodplain
- wetlands
- access
- easements
- environmental overlays

Do not invent the parcel's actual zoning.
""",

        "site": f"""
Analyze site/public works concerns for:

{base_context}

Only identify concerns relevant to the project.
Consider:
- stormwater
- grading
- drainage
- right-of-way
- driveway
- sidewalk
- utilities
- crane
- staging
- dumpster
- public improvements
""",

        "scope": f"""
Analyze scope-specific issues for:

{base_context}

Known concerns:
{chr(10).join("- " + x for x in concerns)}

Use conditional logic.
Explicitly identify what project facts would make each concern
applicable.
""",

        "pitfalls": f"""
Identify practical AHJ coordination pitfalls for:

{base_context}

Focus on:
- permit delays
- missing information
- plan-review corrections
- inspection issues
- jurisdictional confusion
- scope changes
- unanswered questions

Do not present anecdotal practices as official requirements.
""",
    }

    progress = st.progress(0)
    status = st.empty()

    items = list(research_questions.items())

    for index, (key, question) in enumerate(items, start=1):
        status.info(
            f"Research pass {index}/{len(items)}: "
            f"{key.replace('_', ' ').title()}..."
        )

        results[key] = call_ai(question)

        progress.progress(index / len(items))

    status.info("Compiling evidence-aware report...")

    synthesis_prompt = f"""
You are the senior editor compiling an AHJ research package.

Project:
Address: {address}
Project type: {project_type}
Project details: {project_details or "Not provided"}

CRITICAL:
Do not improve certainty merely because multiple AI-generated
research passes agree.

Agreement between AI outputs is NOT independent verification.

Use:

🟢 VERIFIED
Only where a specific authoritative source is cited.

🟡 AI KNOWLEDGE
Likely information without current independent verification.

🟠 UNCERTAIN
Potentially outdated, conditional, incomplete, or conflicting.

🔴 CANNOT VERIFY
Cannot establish.

⚠️ VERIFY
Direct AHJ/property/project confirmation is required.

Never invent a source.

Never invent a code section.

Never invent a zoning designation.

Never invent an AHJ.

Never provide an "Expected Answer."

RESEARCH PASS 1 — JURISDICTION
{results["jurisdiction"]}

RESEARCH PASS 2 — CODES
{results["codes"]}

RESEARCH PASS 3 — PERMITS
{results["permits"]}

RESEARCH PASS 4 — PLANNING
{results["planning"]}

RESEARCH PASS 5 — SITE
{results["site"]}

RESEARCH PASS 6 — SCOPE
{results["scope"]}

RESEARCH PASS 7 — PITFALLS
{results["pitfalls"]}

Create this report:

# AHJ RESEARCH REPORT

## 📊 EXECUTIVE SUMMARY

Give the volunteer:
- overall status
- top 5 unresolved issues
- highest-priority AHJ contacts
- major permit questions

## 📍 JURISDICTION

Separate:
- Established
- Likely
- Needs verification

## 📚 APPLICABLE CODES

For each:
- discipline
- code
- edition
- effective date
- confidence
- source
- applicability
- verification needed

## 📋 PERMIT REQUIREMENTS

Use a table:

| Permit | Status | Trigger | Confidence | What to verify |

Do not say "required" unless evidence supports it.

## 🏘️ PLANNING AND ZONING

Separate actual findings from items that need parcel/AHJ verification.

## 🚜 SITE AND PUBLIC WORKS

Only include issues relevant to this project.

## 🔍 SCOPE-SPECIFIC CONSIDERATIONS

Use conditional logic.

## ⚠️ DIG DEEPER

Rank unresolved issues by:
1. Potential cost impact
2. Potential schedule impact
3. Potential redesign impact

## ❓ QUESTIONS FOR AHJ

Group by department.

For each:
- Exact question
- Why it matters
- Information needed
- Follow-up

Never predict the answer.

## ✅ VOLUNTEER NEXT STEPS

Prioritized checklist.

## 📞 IF INFORMATION IS LIMITED

Tell the volunteer exactly how to obtain the missing information.

## 📚 SOURCES

List only sources actually mentioned in the research.

For each:
- name
- URL
- subject supported
- authority level
- verification status

At the end include:

---
### IMPORTANT
This is a research aid, not an approval or permit determination.
Code editions, permit requirements, zoning, and AHJ procedures must
be confirmed with the applicable authority before relying on them.
---

Return only the report.
"""

    final_report = call_ai(synthesis_prompt)

    progress.empty()
    status.empty()

    return clean_ai_text(final_report)


# ============================================================
# WORD EXPORT
# ============================================================

def markdown_to_docx(text, address, project_type, notes):
    """
    Simple Markdown-to-Word conversion.

    Keeps the report readable without requiring a Markdown parser.
    """

    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    title = document.add_paragraph()
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    run = title.add_run("AHJ RESEARCH REPORT")
    run.bold = True
    run.font.size = Pt(20)

    subtitle = document.add_paragraph()
    subtitle.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    subtitle.add_run(
        f"{project_type}\n{address}\n"
        f"Generated {datetime.datetime.now().strftime('%B %d, %Y')}"
    )

    document.add_paragraph(
        "DRAFT RESEARCH AID — Verify applicable requirements "
        "with the AHJ before relying on this report."
    )

    for raw_line in text.splitlines():
        line = raw_line.strip()

        if not line:
            document.add_paragraph("")
            continue

        if line.startswith("# "):
            p = document.add_heading(line[2:], level=1)
            continue

        if line.startswith("## "):
            document.add_heading(line[3:], level=2)
            continue

        if line.startswith("### "):
            document.add_heading(line[4:], level=3)
            continue

        if line.startswith("- "):
            document.add_paragraph(
                line[2:],
                style="List Bullet",
            )
            continue

        if re.match(r"^\d+\.\s+", line):
            document.add_paragraph(
                re.sub(r"^\d+\.\s+", "", line),
                style="List Number",
            )
            continue

        document.add_paragraph(line)

    if notes:
        document.add_page_break()
        document.add_heading(
            "VOLUNTEER VERIFICATION NOTES",
            level=1,
        )

        document.add_paragraph(
            "These notes record information personally confirmed "
            "by the volunteer with an AHJ or other source."
        )

        for note in notes:
            document.add_paragraph(
                note,
                style="List Bullet",
            )

    document.add_heading("CONFIDENCE GUIDE", level=1)

    guide = [
        "🟢 VERIFIED — Supported by a specific authoritative source.",
        "🟡 AI KNOWLEDGE — Plausible but not independently verified.",
        "🟠 UNCERTAIN — May be outdated, conditional, or incomplete.",
        "🔴 CANNOT VERIFY — Could not establish the answer.",
        "⚠️ VERIFY — Direct AHJ confirmation is required.",
    ]

    for item in guide:
        document.add_paragraph(item)

    document.add_paragraph(
        "IMPORTANT: AI agreement is not independent verification. "
        "Always confirm permit, zoning, code, and jurisdiction "
        "requirements with the applicable authority."
    )

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# SESSION SAVE / LOAD
# ============================================================

def build_session():
    return {
        "app_version": APP_VERSION,
        "saved": now_iso(),
        "address": st.session_state.get("address", ""),
        "project_type": st.session_state.get("project_type", ""),
        "details": st.session_state.get("details", ""),
        "research": st.session_state.get("research", ""),
        "notes": st.session_state.get("notes", []),
        "research_date": st.session_state.get(
            "research_date",
            "",
        ),
    }


def session_download_bytes():
    return json.dumps(
        build_session(),
        indent=2,
        ensure_ascii=False,
    ).encode("utf-8")


def load_session(content):
    try:
        data = json.loads(content)

        if not isinstance(data, dict):
            return None

        if "address" not in data:
            return None

        return data

    except Exception:
        return None


# ============================================================
# INITIAL SESSION STATE
# ============================================================

defaults = {
    "address": "",
    "project_type": "HVAC Replacement",
    "details": "",
    "research": "",
    "notes": [],
    "research_date": "",
    "research_mode": "Deep",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("🏛️ AHJ Research Agent")

    st.caption(
        "Kingdom Hall Project Research Assistant"
    )

    st.markdown("---")

    # --------------------------------------------------------
    # SESSION
    # --------------------------------------------------------

    st.subheader("📂 Resume Session")

    uploaded = st.file_uploader(
        "Load saved session",
        type=["json"],
        key="session_upload",
    )

    if uploaded is not None:

        try:
            content = uploaded.read().decode("utf-8")
            session = load_session(content)

            if session:
                st.session_state["address"] = session.get(
                    "address",
                    "",
                )

                st.session_state["project_type"] = session.get(
                    "project_type",
                    "HVAC Replacement",
                )

                st.session_state["details"] = session.get(
                    "details",
                    "",
                )

                st.session_state["research"] = session.get(
                    "research",
                    "",
                )

                st.session_state["notes"] = session.get(
                    "notes",
                    [],
                )

                st.session_state["research_date"] = session.get(
                    "research_date",
                    "",
                )

                st.success("Session loaded.")

        except Exception as exc:
            st.error(f"Could not load session: {exc}")

    st.markdown("---")

    # --------------------------------------------------------
    # API STATUS
    # --------------------------------------------------------

    st.subheader("🔌 AI Status")

    if api_ready:
        st.success("Gemini API configured")
    else:
        st.error(
            "Gemini API not configured. "
            "Add GEMINI_KEY to Streamlit secrets."
        )

    st.markdown("---")

    # --------------------------------------------------------
    # CONFIDENCE GUIDE
    # --------------------------------------------------------

    st.subheader("📊 Confidence Guide")

    st.markdown(
        """
🟢 **VERIFIED**  
Specific authoritative source supports the finding.

🟡 **AI KNOWLEDGE**  
AI/reference knowledge that has not been independently verified.

🟠 **UNCERTAIN**  
May be outdated, conditional, incomplete, or variable.

🔴 **CANNOT VERIFY**  
The application could not establish the answer.

⚠️ **VERIFY**  
Confirm directly with the AHJ before relying on it.
"""
    )

    st.markdown("---")

    # --------------------------------------------------------
    # IMPORTANT PRINCIPLE
    # --------------------------------------------------------

    st.subheader("⚠️ Important")

    st.info(
        "Multiple AI answers agreeing with one another does NOT "
        "constitute independent verification."
    )

    st.caption(
        f"Version {APP_VERSION} • "
        f"Reference reviewed {REFERENCE_LAST_REVIEWED}"
    )


# ============================================================
# MAIN HEADER
# ============================================================

st.title("🏛️ AHJ Research Agent")

st.caption(
    "Prepare for AHJ conversations — don't replace them."
)


# ============================================================
# PROJECT INPUT
# ============================================================

st.subheader("1. Project Information")

col1, col2 = st.columns([3, 2])

with col1:

    address = st.text_input(
        "Project address",
        value=st.session_state.get("address", ""),
        placeholder="Example: 24340 NW Meek Rd, Hillsboro, OR",
        help=(
            "Use the complete street address whenever possible. "
            "The application will not assume the parcel's exact "
            "jurisdiction unless it can establish it."
        ),
    )

with col2:

    project_type = st.selectbox(
        "Project type",
        list(PROJECT_TYPES.keys()),
        index=(
            list(PROJECT_TYPES.keys()).index(
                st.session_state.get(
                    "project_type",
                    "HVAC Replacement",
                )
            )
        ),
    )


project_info = PROJECT_TYPES[project_type]

st.caption(project_info["description"])


# ============================================================
# SCOPE DETAILS
# ============================================================

st.subheader("2. Scope Details")

st.markdown(
    "The more specific the scope, the better the conditional "
    "research will be."
)

details = st.text_area(
    "Describe the project",
    value=st.session_state.get("details", ""),
    height=180,
    placeholder=(
        "Example:\n"
        "- Existing ground-level HVAC unit\n"
        "- Replace with similar unit\n"
        "- Same location\n"
        "- No ductwork changes\n"
        "- Existing electrical circuit retained\n"
        "- Refrigerant type currently unknown\n"
        "- Building constructed around 1995"
    ),
)


# ============================================================
# SCOPE QUESTIONS
# ============================================================

with st.expander("🔎 Scope questions that may matter"):

    for question in project_info["questions"]:
        st.markdown(f"- {question}")


# ============================================================
# CONCERNS
# ============================================================

selected_concerns = project_info["concerns"]

with st.expander("🧭 Research areas the agent will consider"):

    for concern in selected_concerns:
        st.markdown(f"- {concern}")


# ============================================================
# RESEARCH MODE
# ============================================================

st.subheader("3. Research Depth")

research_mode = st.radio(
    "Choose research depth",
    ["Quick", "Deep"],
    horizontal=True,
    index=1 if st.session_state.get(
        "research_mode",
        "Deep",
    ) == "Deep" else 0,
)

if research_mode == "Quick":
    st.info(
        "Quick mode uses one comprehensive research pass. "
        "Deep mode performs separate research passes before "
        "synthesis and is recommended for real projects."
    )
else:
    st.info(
        "Deep mode separately analyzes jurisdiction, codes, "
        "permits, planning, site issues, scope, and pitfalls "
        "before compiling the final report."
    )


# ============================================================
# RESEARCH BUTTON
# ============================================================

st.subheader("4. Research")

research_button = st.button(
    "🔍 Start AHJ Research",
    type="primary",
    use_container_width=True,
    disabled=not api_ready,
)


if research_button:

    if not address.strip():
        st.error("Enter a project address first.")

    else:

        st.session_state["address"] = address.strip()
        st.session_state["project_type"] = project_type
        st.session_state["details"] = details.strip()
        st.session_state["research_mode"] = research_mode

        with st.spinner(
            "Preparing research..."
        ):

            if research_mode == "Quick":

                result = quick_research(
                    address.strip(),
                    project_type,
                    details.strip(),
                    selected_concerns,
                )

            else:

                result = deep_research(
                    address.strip(),
                    project_type,
                    details.strip(),
                    selected_concerns,
                )

        if result.startswith("ERROR:"):

            st.error(result)

        else:

            st.session_state["research"] = result
            st.session_state["research_date"] = now_iso()

            st.success(
                "Research complete. Review the evidence and "
                "verification questions carefully."
            )

            st.rerun()


# ============================================================
# DISPLAY RESEARCH
# ============================================================

if st.session_state.get("research"):

    st.markdown("---")

    st.subheader("5. Research Report")

    research_date = st.session_state.get(
        "research_date",
        "",
    )

    if research_date:
        st.caption(
            f"Research generated: {research_date}"
        )

    # --------------------------------------------------------
    # Important warning
    # --------------------------------------------------------

    st.warning(
        "This report is a research aid. "
        "AI-generated information is not AHJ confirmation. "
        "Pay particular attention to 🟡, 🟠, 🔴 and ⚠️ findings."
    )

    # --------------------------------------------------------
    # Source diagnostics
    # --------------------------------------------------------

    report_text = st.session_state["research"]
    urls = extract_urls(report_text)

    with st.expander(
        f"🔗 Sources detected in report ({len(urls)})"
    ):

        if not urls:
            st.write(
                "No URLs were detected. "
                "That does not mean the report is wrong; it means "
                "the report does not contain source links that can "
                "be reviewed here."
            )

        else:

            for url in urls:

                official = is_official_domain(url)

                if official:
                    st.markdown(
                        f"🟢 Government-domain source: {url}"
                    )
                else:
                    st.markdown(
                        f"🔗 Source: {url}"
                    )

            st.caption(
                "A government-domain URL is not automatically proof "
                "that the cited page supports the AI's statement. "
                "Open the source and verify the actual content."
            )

    # --------------------------------------------------------
    # Report
    # --------------------------------------------------------

    st.markdown(report_text)

    # ========================================================
    # VOLUNTEER VERIFICATION
    # ========================================================

    st.markdown("---")

    st.subheader("6. Volunteer Verification")

    st.caption(
        "Record what you personally confirm with the AHJ. "
        "These notes are separate from AI-generated findings."
    )

    if "notes" not in st.session_state:
        st.session_state["notes"] = []

    note_col, button_col = st.columns([5, 1])

    with note_col:

        new_note = st.text_input(
            "Verification note",
            placeholder=(
                "Building department confirmed mechanical permit "
                "required; spoke with Jane on 9/10/2026"
            ),
            key="verification_note",
        )

    with button_col:

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(
            "➕ Add",
            key="add_verification",
        ):

            if new_note.strip():

                timestamp = datetime.datetime.now().strftime(
                    "%Y-%m-%d %H:%M"
                )

                st.session_state["notes"].append(
                    {
                        "text": new_note.strip(),
                        "verified": False,
                        "timestamp": timestamp,
                    }
                )

                st.rerun()


    notes = st.session_state["notes"]

    if notes:

        st.markdown("### Verification Log")

        for index, note in enumerate(notes):

            if isinstance(note, str):
                note = {
                    "text": note,
                    "verified": False,
                    "timestamp": "",
                }

            col_a, col_b, col_c = st.columns(
                [8, 1, 1]
            )

            with col_a:

                if note.get("verified"):

                    st.markdown(
                        f"~~☑ {note.get('text', '')}~~"
                    )

                else:

                    st.markdown(
                        f"☐ {note.get('text', '')}"
                    )

                if note.get("timestamp"):
                    st.caption(
                        note["timestamp"]
                    )

            with col_b:

                if not note.get("verified"):

                    if st.button(
                        "✓",
                        key=f"verify_{index}",
                        help="Mark as confirmed",
                    ):

                        st.session_state["notes"][index][
                            "verified"
                        ] = True

                        st.rerun()

            with col_c:

                if st.button(
                    "🗑️",
                    key=f"delete_{index}",
                ):

                    st.session_state["notes"].pop(index)

                    st.rerun()

    else:

        st.info(
            "No volunteer verification notes yet."
        )


    # ========================================================
    # EXPORT
    # ========================================================

    st.markdown("---")

    st.subheader("7. Export")

    export_col1, export_col2, export_col3 = st.columns(3)

    # --------------------------------------------------------
    # Word
    # --------------------------------------------------------

    with export_col1:

        docx_bytes = markdown_to_docx(
            report_text,
            st.session_state["address"],
            st.session_state["project_type"],
            st.session_state["notes"],
        )

        filename_base = re.sub(
            r"[^A-Za-z0-9]+",
            "_",
            st.session_state["project_type"],
        ).strip("_")

        st.download_button(
            "📄 Download Word Report",
            data=docx_bytes,
            file_name=(
                f"AHJ_Research_{filename_base}.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    # --------------------------------------------------------
    # JSON session
    # --------------------------------------------------------

    with export_col2:

        st.download_button(
            "💾 Save Session",
            data=session_download_bytes(),
            file_name="AHJ_Research_Session.json",
            mime="application/json",
            use_container_width=True,
        )

    # --------------------------------------------------------
    # Text
    # --------------------------------------------------------

    with export_col3:

        st.download_button(
            "📝 Download Text",
            data=report_text,
            file_name="AHJ_Research_Report.txt",
            mime="text/plain",
            use_container_width=True,
        )


# ============================================================
# EMPTY STATE
# ============================================================

else:

    st.markdown("---")

    st.info(
        """
### How this version works

**Address + project scope → separate research passes → evidence-aware
synthesis → volunteer verification → Word report**

The application intentionally does **not** assume that an AI answer
is verified simply because the model sounds confident.

For important findings, look for:

🟢 **VERIFIED** — source-supported  
🟡 **AI KNOWLEDGE** — plausible but not independently verified  
🟠 **UNCERTAIN** — conditional or potentially outdated  
🔴 **CANNOT VERIFY** — insufficient evidence  
⚠️ **VERIFY** — ask the AHJ before relying on it
"""
    )

    st.markdown(
        """
### Example project description

> Existing Kingdom Hall. Replace two ground-level HVAC units with
> comparable equipment in the same location. No ductwork changes.
> Existing electrical circuits may remain. Refrigerant type unknown.
> Building constructed approximately 1995.
"""
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    f"AHJ Research Agent v{APP_VERSION} • "
    "Research assistance only • "
    "Verify applicable requirements with the AHJ before relying on them."
)
````
