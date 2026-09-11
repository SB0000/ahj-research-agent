"""
AHJ Research Agent — Kingdom Hall Project Research Assistant
Version 1.0

A research tool for volunteers who serve as liaisons between
Kingdom Hall construction/renovation projects and Authorities
Having Jurisdiction (AHJ).
"""

import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
import json
import datetime
import time

# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="AHJ Research Agent",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# PROJECT KNOWLEDGE BASE
# These are common concerns organized by project type.
# They help the AI ask better scope-specific questions.
# ============================================================

PROJECT_TYPES = {
    "Reroof": {
        "description": "Reroof/replacement of existing roof system",
        "common_concerns": [
            "Like-for-like vs. change of roof covering type",
            "Insulation removal/replacement triggers energy code",
            "Roof deck replacement triggers structural review",
            "Rooftop equipment relocation during reroof",
            "Crane permits and staging area requirements",
            "Dumpster/material staging in public right-of-way",
            "Fall protection requirements during construction",
            "Some jurisdictions exempt like-for-like reroof permits",
            "Wind/uplift requirements may differ from original installation",
            "Tear-off vs. recover changes scope significantly",
        ]
    },
    "HVAC Replacement": {
        "description": "Replacement of HVAC equipment",
        "common_concerns": [
            "Mechanical permit almost always required",
            "Electrical permit if new circuits or disconnects needed",
            "Energy code compliance for new equipment efficiency",
            "Structural review if equipment weight or location changes",
            "Refrigerant type and environmental requirements",
            "Noise ordinance compliance for rooftop equipment",
            "Screening or enclosure requirements for equipment",
            "Roof penetration and flashing details",
            "Ductwork modifications may trigger fire damper requirements",
            "Load calculations may be required",
        ]
    },
    "Parking Lot Improvements": {
        "description": "Parking lot modification, expansion, or improvement",
        "common_concerns": [
            "Stormwater management almost always triggered by impervious area",
            "ADA accessibility upgrades likely required for altered areas",
            "Lighting requirements (poles, photometrics, dark sky)",
            "Landscaping and screening requirements",
            "Driveway access and sight distance requirements",
            "Right-of-way improvements may be triggered",
            "Grading and drainage plan required",
            "Striping, signage, and pavement marking requirements",
            "Parking count reduction may trigger planning review",
            "Infiltration and LID requirements increasing nationally",
        ]
    },
    "Interior Remodel": {
        "description": "Interior remodel with no structural changes",
        "common_concerns": [
            "ADA accessibility triggered by alteration (proportional)",
            "Fire and life safety upgrades may be required",
            "Occupancy classification review if layout changes",
            "Egress and path of travel requirements",
            "Energy code for any envelope work",
            "Asbestos survey if building was built before 1980s",
            "Mechanical permit if HVAC zones change",
            "Electrical permit if circuits are modified",
            "Plumbing permit if fixtures change",
            "Fire sprinkler modifications may require fire dept review",
        ]
    },
    "Major Remodel": {
        "description": "Major remodel, possible structural changes",
        "common_concerns": [
            "All interior remodel concerns plus the following:",
            "Structural review required for wall removal or openings",
            "Change of occupancy possible — triggers full compliance",
            "Full ADA compliance may be triggered depending on scope",
            "Fire sprinkler system may be required or need modification",
            "Fire alarm system upgrades",
            "Energy code compliance for building envelope",
            "Potentially triggers full code compliance review",
            "Value of improvements may trigger proportional upgrades",
        ]
    },
    "Addition": {
        "description": "Building addition or expansion",
        "common_concerns": [
            "Planning and zoning review almost always required",
            "Setback and height compliance for new portion",
            "Parking count impact from increased area",
            "Stormwater for increased impervious area",
            "Fire separation between existing and new construction",
            "Structural connection details to existing building",
            "Full ADA compliance for new portion",
            "Energy code compliance for new building envelope",
            "All applicable permits for new construction scope",
            "Utility capacity for increased load",
        ]
    },
    "New Construction": {
        "description": "New Kingdom Hall construction on a parcel",
        "common_concerns": [
            "Land use approval first — this is the critical path",
            "Site development plan required",
            "All building permits required",
            "Fire and life safety system requirements",
            "Full ADA compliance",
            "Utility connections and availability confirmation",
            "Stormwater management plan",
            "Traffic and access study potentially required",
            "Environmental review (SEPA/NEPA) may apply",
            "Impact fees possible for new development",
            "Pre-application meeting strongly recommended",
        ]
    },
    "Change of Use": {
        "description": "Change of occupancy or use classification",
        "common_concerns": [
            "Change of occupancy can trigger full code compliance",
            "Existing building evaluation required",
            "Fire and life safety upgrades almost certain",
            "Accessibility upgrades required to current standards",
            "Parking recalculation for new use classification",
            "Structural evaluation for new occupancy load",
            "Energy code compliance for new use",
            "Can be as complex as new construction",
            "Certificate of occupancy required for new use",
        ]
    },
    "Site Work Only": {
        "description": "Grading, utilities, or site improvements only",
        "common_concerns": [
            "Grading permit requirements and thresholds",
            "Erosion control plan required",
            "Stormwater management for disturbed area",
            "Utility connection permits and fees",
            "Driveway and access permits",
            "Right-of-way improvements may be triggered",
            "Tree removal and landscaping requirements",
            "Environmental constraints and buffers",
            "Floodplain considerations if applicable",
        ]
    },
}

# ============================================================
# API CONFIGURATION
# ============================================================

@st.cache_resource
def configure_api():
    """Configure the Gemini API — runs once per app lifecycle"""
    try:
        genai.configure(api_key=st.secrets["GEMINI_KEY"])
        return True
    except Exception as e:
        st.error(f"API configuration error: {e}")
        return False

api_ready = configure_api()


def call_ai(prompt, max_retries=2):
    """
    Call Gemini API with error handling and basic retry logic.
    Uses gemini-2.0-flash for speed and cost efficiency.
    """
    if not api_ready:
        return "Error: API not configured. Check your GEMINI_KEY in secrets."

        model = genai.GenerativeModel("gemini-3.6-flash")

    for attempt in range(max_retries + 1):
        try:
            response = model.generate_content(prompt)
            if response.text:
                return response.text
            else:
                return (
                    "Error: No response received from AI. "
                    "This sometimes happens with sensitive topics. "
                    "Please rephrase and try again."
                )
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str:
                if attempt < max_retries:
                    wait_time = 10 * (attempt + 1)
                    time.sleep(wait_time)
                    continue
                return (
                    "Rate limit reached. The free tier allows "
                    "15 requests per minute. Please wait a minute "
                    "and try again."
                )
            elif "403" in error_str:
                return (
                    "API key error. Please check that your Gemini "
                    "API key is valid and properly configured."
                )
            elif "400" in error_str:
                return (
                    "Request error. The prompt may be too long or "
                    "contain unsupported content. Try simplifying "
                    "your project details."
                )
            else:
                if attempt < max_retries:
                    time.sleep(3)
                    continue
                return f"Error after {max_retries + 1} attempts: {error_str}"

    return "Error: Unable to get a response. Please try again."


# ============================================================
# RESEARCH PROMPTS
# ============================================================

def build_research_prompt(address, project_type, project_details, scope_concerns):
    """
    Build the comprehensive research prompt.
    This is the core of the application — a well-structured prompt
    produces well-structured, useful output.
    """

    concerns_text = "\n".join(f"  - {c}" for c in scope_concerns)

    prompt = f"""You are an expert AHJ (Authority Having Jurisdiction) research \
assistant specializing in Kingdom Hall (religious assembly) construction and \
renovation projects. You help volunteers who serve as the official liaison \
between the AHJ and the project team.

Your job is to produce a practical, actionable research report — NOT a \
comprehensive legal analysis. Volunteers need to know what to research, \
who to contact, and what to ask. They will VERIFY everything with the AHJ.

PROJECT INFORMATION:
- Address: {address}
- Project Type: {project_type}
- Building Type: Kingdom Hall (religious assembly, typically A-3 occupancy)
- Scope Details: {project_details if project_details else 'Not provided — assume typical scope'}

COMMON CONCERNS FOR THIS PROJECT TYPE (consider each of these):
{concerns_text}

CONFIDENCE INDICATORS — Use these throughout your response:
  🟢 CONFIRMED  — Found in reliable/official source
  🟡 LIKELY     — Standard practice, but verify locally
  🟠 UNCERTAIN  — Could not confirm, check with AHJ
  🔴 UNKNOWN    — No information found, jurisdiction may lack online resources
  🔄 VARIABLE   — Depends heavily on local conditions or interpretation

RESEARCH AND PROVIDE A STRUCTURED REPORT USING THESE HEADERS:

## 📍 JURISDICTION
Identify the AHJ structure for this location:
- Primary building department (city or county — specify which)
- Planning and zoning jurisdiction (may differ from building)
- Fire authority or fire marshal
- Public works department (if right-of-way affected)
- Health department (if relevant to project scope)
- Any special districts (water, sewer, fire protection, etc.)
- Note: For unincorporated areas, the county is typically the AHJ

For each, indicate confidence level. If you are UNSURE about the \
jurisdiction structure, flag it clearly.

## 📚 APPLICABLE CODES
For EACH of these code categories, provide:
  - Full code name and edition year
  - Whether state-adopted, locally-adopted, or model code referenced
  - Effective date if known
  - Confidence indicator with brief explanation
  - If UNKNOWN: explain why and suggest how to find out

Categories (cover ALL of these):
  1. Building Code
  2. Mechanical Code
  3. Plumbing Code
  4. Electrical Code
  5. Energy Code
  6. Fire/Life Safety Code
  7. Accessibility Code (ADA and state equivalents)
  8. Existing Building Code (if alteration/renovation)

## 📋 PERMIT REQUIREMENTS
Evaluate EACH of these potential permits:
  Building, Mechanical, Plumbing, Electrical, Roofing,
  Grading, Site Development, Stormwater/Drainage,
  Right-of-Way / Encroachment, Driveway Access,
  Planning / Land Use Application, Fire System,
  Sign, Demolition (if applicable), Other (specify)

For each permit provide:
  - Status: [Required | Likely | Possible | Unlikely | Not Required]
  - Confidence: [🟢 | 🟡 | 🟠 | 🔴 | 🔄]
  - Brief reason
  - ⚠️ DIG DEEPER flag if more investigation needed (explain why)

## 🏘️ PLANNING AND ZONING
Evaluate for this specific address and project:
  - Zoning designation (if known or can be inferred)
  - Whether religious assembly is: permitted / conditional / special use
  - Review triggers: site plan, design review, conditional use, etc.
  - Overlay districts: historic, design, flood, environmental
  - Parking requirements and whether this project affects parking count
  - Setback, height, and lot coverage considerations
  - Landscaping requirements
  - Any known special provisions for religious institutions
  For each: confidence level and whether to dig deeper

## 🚜 SITE AND PUBLIC WORKS
Evaluate for this project scope:
  - Right-of-way impacts (construction staging, dumpster, crane)
  - Encroachment permit needs
  - Stormwater and drainage requirements
  - Utility considerations and connection permits
  - Access and driveway requirements
  - Sidewalk, curb, and street improvement requirements
  - Erosion and sediment control
  - Tree protection or removal requirements
  For each: status, confidence, and dig deeper flags

## 🔍 SCOPE-SPECIFIC CONSIDERATIONS
Based on the EXACT project scope described above:
  - Which specific code sections most likely apply to THIS scope
  - What aspects of this scope might trigger ADDITIONAL requirements \
not obvious at first
  - What exemptions might apply (religious assembly, like-for-like, \
minor work, etc.)
  - What common pitfalls arise with this exact type of project
  - What the volunteer should specifically watch out for
  - What details about the scope matter for permit requirements
  For each: practical explanation and what to verify

## ⚠️ DIG DEEPER
List areas that need additional research beyond what you can provide. \
For each flag:
  - Topic area
  - WHY you should dig deeper (be specific)
  - What exactly to investigate
  - Suggested approach: call AHJ, check website, ask local contractor, \
check with neighboring jurisdiction, etc.

Common reasons to flag:
  - Jurisdiction appears to have limited online presence
  - Local amendments could significantly modify state requirements
  - Answer depends on specific project scope details not yet known
  - Religious assemblies may have special provisions or exemptions
  - Historic districts or overlay zones may apply but cannot confirm
  - Older existing building may have pre-existing non-conformities
  - Rural or small jurisdiction — limited resources or information
  - This requirement varies significantly between jurisdictions
  - The AHJ may have unwritten practices that differ from published rules

## ❓ QUESTIONS FOR AHJ
Generate specific questions organized by AHJ department or topic. \
For EACH question:
  - The exact question to ask
  - Why it matters (what could go wrong if you don't ask)
  - What answer you are hoping for or expect
  - Follow-up question if the answer is unexpected

Do NOT ask obvious questions the AHJ website probably answers. \
Ask the questions that are AMBIGUOUS, VARIABLE, or NOT FINDABLE online.

## ✅ VOLUNTEER NEXT STEPS
Provide a numbered, priority-ordered, ACTION-ORIENTED list.
Tell the volunteer WHAT TO DO — not just what the code says.
Include: who to contact, what to ask, what to prepare, \
what order to do things in, and what decisions need to be made first.

Start with the most critical/urgent item.

## 📞 IF INFORMATION IS LIMITED
If this jurisdiction appears to have limited online resources:
  - Acknowledge this clearly
  - Suggest how to proceed when you can't find information online
  - What to say when calling the AHJ
  - What specific items to ask for when you reach someone
  - Whether a pre-application meeting is recommended
  - Alternative sources: neighboring contractors, regional building \
council, state building code division, etc.

CRITICAL RULES:
- Never present AI inference as confirmed fact
- When uncertain, flag it clearly and explain what to verify
- Religious assemblies (A-3 occupancy) sometimes have special provisions — mention this
- Small and rural jurisdictions often have limited online presence — acknowledge this
- State vs. local code adoption varies significantly — be specific about sources
- Always include verification guidance
- Be PRACTICAL and ACTIONABLE, not theoretical
- Focus on what the volunteer NEEDS TO DO, not just what the code says
"""
    return prompt


# ============================================================
# RESEARCH FUNCTIONS
# ============================================================

def quick_research(address, project_type, project_details, scope_concerns):
    """Single comprehensive query — faster but less thorough"""
    prompt = build_research_prompt(
        address, project_type, project_details, scope_concerns
    )
    with st.spinner("🔍 Researching jurisdiction and requirements..."):
        return call_ai(prompt)


def deep_research(address, project_type, project_details, scope_concerns):
    """
    Multiple targeted queries for more thorough research.
    Each query focuses on a different aspect, then we synthesize.
    This produces better results for complex or unfamiliar jurisdictions.
    """
    results = {}
    project_desc = PROJECT_TYPES.get(project_type, {}).get("description", project_type)
    concerns_list = "\n".join(f"- {c}" for c in scope_concerns)

    # --- Query 1: Jurisdiction & Codes ---
    with st.spinner("🔍 Step 1/6: Identifying jurisdiction and adopted codes..."):
        q1 = f"""What building, mechanical, plumbing, electrical, energy, and fire \
codes are currently adopted and effective for construction projects at \
{address}?

Include state-adopted codes and any known local amendments or modifications.
For each code, provide: full name, edition year, and effective date.
If this is a small or rural jurisdiction with limited online presence, say so.

For each code, use a confidence indicator:
  🟢 CONFIRMED | 🟡 LIKELY | 🟠 UNCERTAIN | 🔴 UNKNOWN | 🔄 VARIABLE

Also identify the jurisdiction structure: is the building department city-run, \
county-run, or a combined city-county agency? Who handles planning vs. building?"""
        results["codes"] = call_ai(q1)

    # --- Query 2: Permits ---
    with st.spinner("🔍 Step 2/6: Researching permit requirements..."):
        q2 = f"""What permits are typically required for a {project_desc} \
at an existing Kingdom Hall (religious assembly, A-3 occupancy) at {address}?

Evaluate each potential permit:
Building, Mechanical, Plumbing, Electrical, Roofing, Grading, \
Site Development, Stormwater, Right-of-Way/Encroachment, \
Driveway Access, Planning/Land Use, Fire, Sign, Demolition

For each: status (Required/Likely/Possible/Unlikely/Not Required), \
confidence (🟢🟡🟠🔴🔄), and reason.

Flag any ⚠️ DIG DEEPER items with explanation of why more research is needed.

Consider whether religious assembly status affects any permit requirements."""
        results["permits"] = call_ai(q2)

    # --- Query 3: Planning & Zoning ---
    with st.spinner("🔍 Step 3/6: Checking planning and zoning requirements..."):
        q3 = f"""What planning and zoning requirements apply to a \
{project_desc} at {address}?

Consider:
- Zoning designation and whether religious assembly is a permitted use
- Conditional use or special use permit requirements for religious assembly
- Site plan review triggers
- Design review or historic district overlay
- Parking requirements — does this project change the parking count?
- Setback, height, and lot coverage
- Overlay districts (historic, design, flood, environmental)
- Landscaping requirements
- Any special provisions for religious institutions in this jurisdiction

Mark items needing verification. Flag dig deeper areas with ⚠️."""
        results["planning"] = call_ai(q3)

    # --- Query 4: Site & Public Works ---
    with st.spinner("🔍 Step 4/6: Evaluating site and public works..."):
        q4 = f"""What site work and public works requirements might apply to a \
{project_desc} at {address}?

Consider:
- Right-of-way impacts (construction staging, dumpster placement, crane operation)
- Encroachment permits for work in public right-of-way
- Stormwater management and drainage requirements
- Utility connections, permits, and availability
- Driveway access and sight distance requirements
- Sidewalk, curb, and street improvement requirements
- Erosion and sediment control
- Tree protection or removal permits
- Grading permit thresholds

Mark items needing verification. Flag dig deeper areas with ⚠️."""
        results["site"] = call_ai(q4)

    # --- Query 5: Scope-Specific ---
    with st.spinner("🔍 Step 5/6: Analyzing scope-specific concerns..."):
        q5 = f"""For a {project_desc} at a Kingdom Hall at {address}, \
what scope-specific concerns should a volunteer investigate?

Common concerns for this project type:
{concerns_list}

Additional project details: {project_details if project_details else 'Not provided'}

For each concern:
- Explain WHY it matters in practical terms
- What could go wrong if it is not addressed
- What specifically to verify with the AHJ
- What questions to ask

Also identify any special provisions, exemptions, or considerations that \
might apply to religious assemblies (A-3 occupancy) for this specific \
type of work in this jurisdiction."""
        results["scope"] = call_ai(q5)

    # --- Query 6: Common Pitfalls ---
    with st.spinner("🔍 Step 6/6: Researching common pitfalls and local knowledge..."):
        q6 = f"""What are common pitfalls, complications, or issues that volunteers \
should be aware of when dealing with AHJs for a {project_desc} at a Kingdom \
Hall (religious assembly) at {address}?

Consider:
- Common mistakes volunteers make with this project type
- Things AHJs frequently ask about or require for Kingdom Halls specifically
- Local practices that may differ from published requirements
- Items that commonly delay permit issuance or plan review
- Inspection requirements that catch people off guard
- Requirements that are frequently overlooked
- Things that seem simple but become complex with AHJ review
- Seasonal or timing considerations (review backlog, inspection scheduling)
- Common reasons for corrections or re-submittal

Be practical and specific. This is for volunteers who may not be \
construction professionals but serve as the AHJ liaison."""
        results["pitfalls"] = call_ai(q6)

    # --- Synthesize ---
    with st.spinner("📝 Compiling comprehensive research report..."):
        synth_prompt = f"""You are compiling a comprehensive AHJ research \
report for a Kingdom Hall {project_type} project at {address}.

Project details: {project_details if project_details else 'Not provided'}

RESEARCH RESULTS TO SYNTHESIZE:

=== JURISDICTION AND CODES ===
{results['codes']}

=== PERMITS ===
{results['permits']}

=== PLANNING AND ZONING ===
{results['planning']}

=== SITE AND PUBLIC WORKS ===
{results['site']}

=== SCOPE-SPECIFIC CONCERNS ===
{results['scope']}

=== COMMON PITFALLS ===
{results['pitfalls']}

Compile this into a SINGLE well-organized report using EXACTLY these headers:

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

Rules for synthesis:
- Organize and consolidate — do NOT just concatenate the raw research
- Resolve contradictions if found (note them if unresolvable)
- Mark everything uncertain with the appropriate confidence indicator
- Flag all dig deeper items with ⚠️
- Questions for AHJ should be specific, practical, and actionable
- Next steps should tell the volunteer EXACTLY what to do in priority order
- Never present AI inference as confirmed fact
- If this jurisdiction has limited online presence, acknowledge it prominently
- Include specific guidance for how to proceed when information is limited
- Make the report USEFUL for a volunteer who needs to make phone calls tomorrow
"""
        return call_ai(synth_prompt)


# ============================================================
# WORD DOCUMENT EXPORT
# ============================================================

def create_word_report(address, project_type, research_text, notes=None):
    """Create a professional Word document from the research results"""

    doc = Document()

    # --- Title ---
    title = doc.add_heading("AHJ Research Report", level=0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # --- Project Information ---
    info = doc.add_paragraph()
    run_project = info.add_run("Project Type: ")
    run_project.bold = True
    info.add_run(f"{project_type}\n")
    run_addr = info.add_run("Address: ")
    run_addr.bold = True
    info.add_run(f"{address}\n")
    run_date = info.add_run("Research Date: ")
    run_date.bold = True
    info.add_run(f'{datetime.datetime.now().strftime("%B %d, %Y")}\n')
    run_status = info.add_run("Status: ")
    run_status.bold = True
    info.add_run("DRAFT — Verify all items with AHJ before relying on this information")

    doc.add_paragraph("─" * 60)

    # --- Parse and Add Research Content ---
    for line in research_text.split("\n"):
        stripped = line.strip()

        if not stripped:
            continue

        # Headers
        if stripped.startswith("## "):
            heading_text = stripped.replace("## ", "").strip()
            doc.add_heading(heading_text, level=1)
        elif stripped.startswith("### "):
            heading_text = stripped.replace("### ", "").strip()
            doc.add_heading(heading_text, level=2)
        # Bullet points
        elif stripped.startswith("- ") or stripped.startswith("  - "):
            bullet_text = stripped.lstrip(" -")
            doc.add_paragraph(bullet_text, style="List Bullet")
        # Checkbox items
        elif stripped.startswith("☐") or stripped.startswith("☑"):
            doc.add_paragraph(stripped, style="List Bullet")
        # Numbered items
        elif len(stripped) > 2 and stripped[0].isdigit() and stripped[1] in ".":
            doc.add_paragraph(stripped, style="List Number")
        # Regular text (handle bold markers)
        elif "**" in stripped:
            para = doc.add_paragraph()
            parts = stripped.split("**")
            for i, part in enumerate(parts):
                if i % 2 == 1:
                    para.add_run(part).bold = True
                else:
                    para.add_run(part)
        else:
            doc.add_paragraph(stripped)

    # --- Verification Notes ---
    if notes:
        doc.add_heading("Volunteer Verification Notes", level=1)
        doc.add_paragraph(
            "Record confirmations, corrections, and additional information "
            "from AHJ contact:"
        )
        for note in notes:
            doc.add_paragraph(note, style="List Bullet")

    # --- Disclaimer Footer ---
    doc.add_paragraph("─" * 60)
    footer = doc.add_paragraph()
    run_disc = footer.add_run("⚠️ DISCLAIMER: ")
    run_disc.bold = True
    footer.add_run(
        "This report is AI-generated research for volunteer use only. "
        "All information must be verified with the Authority Having "
        "Jurisdiction (AHJ) before any reliance is placed on it. "
        "Building codes, permit requirements, and local amendments "
        "change frequently and vary by jurisdiction. AI knowledge has "
        "a cutoff date and may not reflect the most current adoptions. "
        "This report does not constitute legal, engineering, or "
        "professional advice."
    )

    # --- Save to Buffer ---
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# SESSION SAVE AND LOAD
# ============================================================

def save_session(address, project_type, details, research, notes):
    """Save current session to a JSON string"""
    session = {
        "address": address,
        "project_type": project_type,
        "details": details,
        "research": research,
        "notes": notes,
        "saved": datetime.datetime.now().isoformat(),
        "version": "1.0",
    }
    return json.dumps(session, indent=2)


def load_session(json_string):
    """Load a session from JSON string"""
    try:
        return json.loads(json_string)
    except json.JSONDecodeError:
        return None


# ============================================================
# USER INTERFACE
# ============================================================

# --- Sidebar ---
with st.sidebar:
    st.header("🏛️ AHJ Research Agent")
    st.caption("Kingdom Hall Project Research Assistant")
    st.markdown("---")

    # Resume session
    st.subheader("📂 Resume Session")
    uploaded = st.file_uploader(
        "Load a saved session (.json)",
        type=["json"],
        key="session_upload",
        help="Upload a .json file you saved from a previous session",
    )

    if uploaded is not None:
        try:
            content = uploaded.read().decode("utf-8")
            session = load_session(content)
            if session and "address" in session:
                st.session_state["address"] = session.get("address", "")
                st.session_state["project_type"] = session.get("project_type", "")
                st.session_state["details"] = session.get("details", "")
                st.session_state["research"] = session.get("research", "")
                st.session_state["notes"] = session.get("notes", [])
                saved_date = session.get("saved", "")[:10]
                st.success(f"✅ Loaded session from {saved_date}")
                st.rerun()
            else:
                st.error("Invalid session file format")
        except Exception as e:
            st.error(f"Error loading session: {e}")

    st.markdown("---")

    # How it works
    st.subheader("ℹ️ How It Works")
    st.markdown(
        """
1. Enter the project address
2. Select the project type
3. Add any scope details
4. Click **Research**
5. Review the findings
6. Add verification notes as you contact the AHJ
7. Export to Word or save session

**Critical:** All findings must be verified with the AHJ. This tool accelerates research — it does not replace AHJ contact.
        """
    )

    st.markdown("---")

    # Confidence guide
    st.subheader("📊 Confidence Guide")
    st.markdown(
        """
🟢 **CONFIRMED** — Found in official source

🟡 **LIKELY** — Standard practice, verify locally

🟠 **UNCERTAIN** — Could not confirm, check with AHJ

🔴 **UNKNOWN** — No info found, jurisdiction may lack online resources

🔄 **VARIABLE** — Depends on local conditions
        """
    )

    st.markdown("---")

    # Feature list
    st.subheader("✨ Features in v1.0")
    st.markdown(
        """
✅ Jurisdiction identification
✅ Code cycles with confidence
✅ Permit requirements analysis
✅ Planning & zoning review
✅ Site & public works
✅ Scope-specific concerns
✅ Dig deeper flags
✅ AHJ questions by department
✅ Action-oriented next steps
✅ Limited-info guidance
✅ Verification checklist
✅ Word document export
✅ Session save & resume
✅ Deep research mode
✅ Project knowledge base
        """
    )

    st.markdown("---")
    st.caption("v1.0 — Use at your own risk. Verify everything with AHJ.")


# --- Main Content ---

st.title("🏛️ AHJ Research Agent")
st.caption(
    "Kingdom Hall Project Research Assistant — "
    "All findings must be verified with the AHJ"
)

# --- Input Form ---
with st.container():
    col1, col2 = st.columns([3, 2])

    with col1:
        address = st.text_input(
            "📍 Project Address",
            value=st.session_state.get("address", ""),
            placeholder="123 Main St, Springfield, OR 97477",
            help="Enter the full address including city and state",
        )

    with col2:
        # Determine default project type from saved session
        saved_type = st.session_state.get("project_type", "")
        type_keys = list(PROJECT_TYPES.keys())
        default_index = type_keys.index(saved_type) if saved_type in type_keys else 0

        project_type = st.selectbox(
            "🔨 Project Type",
            type_keys,
            index=default_index,
            help="Select the type of project",
        )

    details = st.text_area(
        "📝 Project Scope Details (optional but recommended)",
        value=st.session_state.get("details", ""),
        placeholder=(
            "Examples:\n"
            "- Existing Kingdom Hall, approx 8,000 sq ft, flat roof, single story\n"
            "- Replace built-up roof with TPO membrane, same insulation thickness\n"
            "- No change to rooftop equipment or roof deck\n"
            "- Dumpster will be in parking lot, not in right-of-way\n"
            "- Building was originally constructed in 1992"
        ),
        height=130,
        help="More details = better research. Include building info, scope specifics, and constraints.",
    )

    # Show common concerns for selected project type
    if project_type in PROJECT_TYPES:
        with st.expander(f"📋 Common concerns for {project_type}", expanded=False):
            for concern in PROJECT_TYPES[project_type]["common_concerns"]:
                st.markdown(f"- {concern}")

    # Research depth selection
    col_depth1, col_depth2 = st.columns([1, 3])
    with col_depth1:
        research_depth = st.radio(
            "Research depth:",
            ["Quick", "Deep"],
            horizontal=True,
            help="Deep mode runs 6 targeted queries then synthesizes — better for new or complex projects",
        )

    with col_depth2:
        st.markdown(
            "**Quick:** 1 query (~20-30 sec) — good for simple projects &nbsp;|&nbsp; "
            "**Deep:** 6 queries + synthesis (~2-3 min) — recommended for new or complex projects"
        )

# --- Research Button ---
st.markdown("---")

if st.button("🔍 Research AHJ Requirements", type="primary", use_container_width=True):
    if not address:
        st.error("⚠️ Please enter a project address")
    elif len(address.strip()) < 10:
        st.warning(
            "⚠️ Please enter a more complete address — "
            "include city and state for best results"
        )
    elif not api_ready:
        st.error("⚠️ API not configured. Check the GEMINI_KEY in settings.")
    else:
        # Get scope concerns for this project type
        scope_concerns = PROJECT_TYPES.get(project_type, {}).get(
            "common_concerns", []
        )

        # Run research
        if research_depth == "Quick":
            result = quick_research(address, project_type, details, scope_concerns)
        else:
            result = deep_research(address, project_type, details, scope_concerns)

        # Store results in session state
        st.session_state["research"] = result
        st.session_state["address"] = address
        st.session_state["project_type"] = project_type
        st.session_state["details"] = details
        st.session_state["research_date"] = datetime.datetime.now().isoformat()

        if result.startswith("Error"):
            st.error(result)
        else:
            st.success("✅ Research complete — review findings below")


# --- Display Results ---
if "research" in st.session_state and st.session_state["research"]:

    # Show research timestamp
    research_date = st.session_state.get("research_date", "")
    if research_date:
        display_date = research_date[:19].replace("T", " at ")
        st.caption(f"📅 Research completed: {display_date}")

    # Main research display
    st.markdown(st.session_state["research"])

    # ============================================================
    # VERIFICATION SECTION
    # ============================================================
    st.markdown("---")
    st.subheader("✅ Volunteer Verification Notes")
    st.caption(
        "Add notes as you confirm items with the AHJ. "
        "These will be included in your exported report."
    )

    # Initialize notes list
    if "notes" not in st.session_state:
        st.session_state["notes"] = []

    notes = st.session_state["notes"]

    # Add new note
    col_new1, col_new2 = st.columns([5, 1])
    with col_new1:
        new_note = st.text_input(
            "Add a verification note:",
            placeholder="Called building dept — confirmed reroof permit required, no planning review needed",
            key="new_note_input",
        )
    with col_new2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Add Note", key="add_note_btn"):
            if new_note.strip():
                timestamp = datetime.datetime.now().strftime("%m/%d %H:%M")
                notes.append(f"☐ {new_note.strip()} ({timestamp})")
                st.session_state["notes"] = notes
                st.rerun()

    # Display existing notes with check-off and delete
    if notes:
        st.markdown("---")
        for i, note in enumerate(notes):
            col_note, col_check, col_del = st.columns([10, 1, 1])
            with col_note:
                if note.startswith("☑"):
                    st.markdown(f"~~{note}~~")
                else:
                    st.markdown(note)
            with col_check:
                if note.startswith("☐"):
                    if st.button("✓", key=f"check_{i}", help="Mark as verified"):
                        notes[i] = notes[i].replace("☐", "☑", 1)
                        st.session_state["notes"] = notes
                        st.rerun()
            with col_del:
                if st.button("🗑️", key=f"del_{i}", help="Delete this note"):
                    notes.pop(i)
                    st.session_state["notes"] = notes
                    st.rerun()
    else:
        st.info(
            "No verification notes yet. Add notes as you contact the AHJ to confirm findings."
        )

    # ============================================================
    # EXPORT SECTION
    # ============================================================
    st.markdown("---")
    st.subheader("📥 Export & Save")

    col_x, col_y, col_z = st.columns(3)

    with col_x:
        try:
            doc_buf = create_word_report(
                st.session_state.get("address", ""),
                st.session_state.get("project_type", ""),
                st.session_state["research"],
                notes,
            )
            # Create safe filename
            safe_name = st.session_state.get("address", "report").split(",")[0]
            safe_name = "".join(
                c for c in safe_name if c.isalnum() or c in " _-"
            ).strip()[:30]
            date_str = datetime.datetime.now().strftime("%Y%m%d")

            st.download_button(
                "📄 Export to Word",
                data=doc_buf.getvalue(),
                file_name=f"AHJ_Research_{safe_name}_{date_str}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                help="Download a Word document for your project file",
            )
        except Exception as e:
            st.error(f"Word export error: {e}")

    with col_y:
        try:
            session_json = save_session(
                st.session_state.get("address", ""),
                st.session_state.get("project_type", ""),
                st.session_state.get("details", ""),
                st.session_state["research"],
                notes,
            )
            date_str = datetime.datetime.now().strftime("%Y%m%d")

            st.download_button(
                "💾 Save Session",
                data=session_json,
                file_name=f"AHJ_session_{date_str}.json",
                mime="application/json",
                use_container_width=True,
                help="Save to resume later — upload in the sidebar next time",
            )
        except Exception as e:
            st.error(f"Session save error: {e}")

    with col_z:
        # Plain text export
        research_text = st.session_state["research"]
        if notes:
            research_text += "\n\n## Volunteer Verification Notes\n"
            research_text += "\n".join(f"- {n}" for n in notes)
        date_str = datetime.datetime.now().strftime("%Y%m%d")

        st.download_button(
            "📋 Export as Text",
            data=research_text,
            file_name=f"AHJ_Research_{date_str}.txt",
            mime="text/plain",
            use_container_width=True,
            help="Download as plain text",
        )

    # --- Disclaimer ---
    st.markdown("---")
    st.warning(
        """
⚠️ **DISCLAIMER:** This report is AI-generated research for volunteer use only.
All information must be verified with the Authority Having Jurisdiction (AHJ)
before any reliance is placed on it. Building codes, permit requirements, and
local amendments change frequently. AI knowledge has a cutoff date and may not
reflect the most current code adoptions. This report does not constitute legal,
engineering, or professional advice.
        """
    )
