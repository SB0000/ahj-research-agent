import os
import re
import json
import urllib.request
import urllib.error
from datetime import datetime, date
from io import BytesIO

import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt

# ============================================================
# AHJ RESEARCH ASSISTANT
# Search-grounded, state-agnostic, official-source-first
#
# IMPORTANT:
# - This version uses the current Google GenAI SDK.
# - Gemini does NOT use Google Search grounding in this quota-safe version.
# - Current web evidence is retrieved separately and supplied to Gemini.
# - Code editions are NOT hard-coded by state.
# - The model must discover the applicable code cycle from current sources.
# ============================================================

st.set_page_config(
    page_title="AHJ Research Assistant",
    page_icon="🏛️",
    layout="wide",
)

# -----------------------------
# State list
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
A specific authoritative source was actually retrieved and supports the statement.

🟡 LIKELY
Strong preliminary conclusion, but exact project/AHJ applicability still needs confirmation.

🟠 CONDITIONAL
True only if a stated condition applies.

🔴 UNKNOWN
Available evidence does not establish the answer.

⚠️ VERIFY
A direct AHJ, parcel record, permit record, code official, or project document check is required.

IMPORTANT:
A source being official does NOT automatically make every claim about that source verified.
A reachable URL does NOT prove that the page supports a claim.
"""

# -----------------------------
# Session state
# -----------------------------
defaults = {
    "report": "",
    "research_log": [],
    "research_sources": [],
    "verified_facts": [],
    "source_checks": [],
    "web_evidence": [],
    "project_fingerprint": {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# -----------------------------
# Secrets / model configuration
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
CONFIGURED_MODEL = secret_or_env("GEMINI_MODEL", "gemini-3.6-flash")

MODEL_NAME = CONFIGURED_MODEL or "gemini-3.6-flash"

# -----------------------------
# Helpers
# -----------------------------
def extract_urls(text):
    if not text:
        return []

    urls = re.findall(r"https?://[^\s\]\)>\"']+", text)
    cleaned = []

    for url in urls:
        url = url.rstrip(".,;:")
        if url not in cleaned:
            cleaned.append(url)

    return cleaned


def check_url(url, timeout=6):
    """
    Reachability check only.
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
            headers={"User-Agent": "AHJ-Research-Assistant/2.0"},
        )
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


def build_fingerprint(
    state,
    project_date,
    project_type,
    classification,
    address,
    details,
    scope_answers,
):
    fp = {
        "state": state,
        "project_date": project_date.isoformat() if isinstance(project_date, date) else str(project_date),
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


def base_system_rules():
    return """
You are an AHJ research assistant for construction volunteers.

You are NOT a generic construction-code chatbot.

Your job is to research the specific project using current web evidence,
prioritize authoritative government sources, identify the few requirements
that actually matter, and clearly distinguish verified evidence from
preliminary interpretation.

============================================================
NON-NEGOTIABLE RESEARCH RULES
============================================================

1. CURRENT WEB RESEARCH IS REQUIRED
For code editions, permit procedures, jurisdiction, zoning, planning,
adopted amendments, fire authority, and other time-sensitive matters,
Use the EXTERNAL WEB EVIDENCE supplied to you. Do not claim to have independently searched the web.

Do not answer these from model memory when current official information
can be searched.

2. OFFICIAL SOURCES FIRST
Prefer, in this order:

A. State government / state building-code agency
B. County / city / municipality / AHJ
C. Official fire authority / fire marshal
D. Official permit portal
E. Official ordinance, adopted code, amendment, interpretation, or policy
F. Official project/land-use record
G. Secondary sources only when authoritative sources are unavailable

3. SOURCE-DISCOVERY IS PART OF THE JOB
If the state or local agency is not already supplied, FIND the correct
official source.

For state code research, search specifically for:
- current adopted building/structural code
- mechanical code
- electrical code
- plumbing code
- fire/life-safety code
- energy code
- existing-building code
- state amendments
- code adoption/effective/mandatory dates
- statewide interpretations or alternate methods when relevant

Do not assume every state uses the same code family, edition, or adoption model.

4. DO NOT HARD-CODE CODE EDITIONS
Never infer that a code is current because it was current in training data.
Never copy the code cycle from another state.
Never assume the International Code Council edition is the state's adopted edition.

The controlling question is:
"What code was actually adopted and applicable to this project in this
jurisdiction on the relevant project/permit date?"

5. PROJECT DATE MATTERS
Code applicability may depend on:
- permit/application date
- effective date
- mandatory date
- phase-in period
- transition provisions
- existing-building provisions
- type of work

Always consider these.

6. EXISTING BUILDING VS NEW WORK
For an existing building, distinguish:
- repair
- alteration
- replacement
- addition
- change of occupancy/use
- change of level of work
- new construction

Do not automatically apply new-construction requirements to a simple
replacement.

7. COMMERCIAL / RESIDENTIAL DISCIPLINE
Do not mix residential and commercial code paths.

The stated building/use classification is a research input, not proof of
the legal occupancy classification. If the legal classification matters,
say so and identify what must be verified.

8. SCOPE DISCIPLINE
Research only issues actually created by the scope.

If the scope explicitly eliminates an issue, do not spend report space
researching it unless there is a reason the elimination may be unreliable.

9. NO INVENTED CODE SECTIONS
Never invent:
- section numbers
- code quotations
- permit thresholds
- zoning designations
- CUP conditions
- fire districts
- inspection requirements
- effective dates
- official URLs

If you cannot verify it, say UNKNOWN or VERIFY.

10. NO FALSE CERTAINTY
"Multiple AI passes agree" is NOT evidence.

"Official-looking URL" is NOT evidence.

A URL being reachable is NOT evidence.

A claim is VERIFIED only when the retrieved source actually supports it.

11. SOURCE CITATIONS
Every material current claim must cite one or more supplied [LIVE SOURCE N] identifiers. If no supplied source supports the claim, mark it UNKNOWN/VERIFY instead of citing model knowledge. Never create a source, URL, or source number.

11. AHJ QUESTIONS
When the evidence does not settle the issue, give the volunteer the
exact question to ask the AHJ rather than guessing the answer.

12. OUTPUT
Be concise and practical. A volunteer should be able to understand the
preliminary result in under two minutes.
"""


def call_gemini(prompt, temperature=0.1):
    """Quota-safe Gemini call. No Gemini Search grounding or URL tools."""
    if not GEMINI_KEY:
        return {
            "text": "ERROR: GEMINI_KEY is not configured.",
            "sources": [],
            "model": "",
            "error": True,
        }

    try:
        client = genai.Client(api_key=GEMINI_KEY)
        config_kwargs = {
            "max_output_tokens": 4000,
        }

        # Gemini 3.x generally performs best with its default sampling
        # settings. Only use temperature for older/non-3.x models.
        if not MODEL_NAME.startswith("gemini-3."):
            config_kwargs["temperature"] = temperature

        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=prompt,
            config=types.GenerateContentConfig(**config_kwargs),
        )

        # The SDK normally exposes response.text, but a response can also
        # contain text parts that are not surfaced by that convenience field.
        text = getattr(response, "text", None)

        if not text:
            parts = []
            candidates = getattr(response, "candidates", None) or []
            for candidate in candidates:
                content = getattr(candidate, "content", None)
                for part in (getattr(content, "parts", None) or []):
                    part_text = getattr(part, "text", None)
                    if part_text:
                        parts.append(part_text)

            if parts:
                text = "\n".join(parts)

        if text:
            return {
                "text": text.strip(),
                "sources": [],
                "model": MODEL_NAME,
                "error": False,
            }

        # Give the UI a useful diagnostic instead of the opaque
        # "no text output" message.
        details = []
        candidates = getattr(response, "candidates", None) or []
        for candidate in candidates:
            finish_reason = getattr(candidate, "finish_reason", None)
            if finish_reason:
                details.append(f"finish_reason={finish_reason}")

        prompt_feedback = getattr(response, "prompt_feedback", None)
        if prompt_feedback:
            details.append(f"prompt_feedback={prompt_feedback}")

        suffix = ("; " + ", ".join(details)) if details else ""
        return {
            "text": (
                f"ERROR: Gemini returned no text output using {MODEL_NAME}"
                f"{suffix}. Check the Gemini API response details in the app logs."
            ),
            "sources": [],
            "model": MODEL_NAME,
            "error": True,
        }

    except Exception as exc:
        return {
            "text": (
                f"ERROR: Gemini request failed using {MODEL_NAME}: "
                f"{type(exc).__name__}: {exc}"
            ),
            "sources": [],
            "model": MODEL_NAME,
            "error": True,
        }


# -----------------------------
# Bounded external web retrieval
# -----------------------------
STATE_SEARCH_HINTS = {
"Alabama":"Alabama Building Commission building codes", "Alaska":"Alaska building codes state", "Arizona":"Arizona building codes state", "Arkansas":"Arkansas building codes state", "California":"California Building Standards Commission building codes", "Colorado":"Colorado building codes state", "Connecticut":"Connecticut State Building Code", "Delaware":"Delaware building code state", "Florida":"Florida Building Code official", "Georgia":"Georgia state minimum standard building code", "Hawaii":"Hawaii building code state", "Idaho":"Idaho building codes state", "Illinois":"Illinois building codes state", "Indiana":"Indiana building codes state", "Iowa":"Iowa building codes state", "Kansas":"Kansas building codes state", "Kentucky":"Kentucky building codes state", "Louisiana":"Louisiana building codes state", "Maine":"Maine building codes state", "Maryland":"Maryland building codes state", "Massachusetts":"Massachusetts building code state", "Michigan":"Michigan building codes state", "Minnesota":"Minnesota building code state", "Mississippi":"Mississippi building codes state", "Missouri":"Missouri building codes state", "Montana":"Montana building codes state", "Nebraska":"Nebraska building codes state", "Nevada":"Nevada building codes state", "New Hampshire":"New Hampshire building code state", "New Jersey":"New Jersey building code state", "New Mexico":"New Mexico building codes state", "New York":"New York State building code", "North Carolina":"North Carolina building code", "North Dakota":"North Dakota building code", "Ohio":"Ohio building code", "Oklahoma":"Oklahoma building code state", "Oregon":"Oregon BCD adopted codes", "Pennsylvania":"Pennsylvania Uniform Construction Code", "Rhode Island":"Rhode Island building code", "South Carolina":"South Carolina building code", "South Dakota":"South Dakota building code", "Tennessee":"Tennessee building code state", "Texas":"Texas building code state", "Utah":"Utah building code state", "Vermont":"Vermont building code state", "Virginia":"Virginia Uniform Statewide Building Code", "Washington":"Washington State Building Code", "West Virginia":"West Virginia building code state", "Wisconsin":"Wisconsin building code state", "Wyoming":"Wyoming building code state", "District of Columbia":"District of Columbia building code"}

def _clean_html_text(raw):
    from html.parser import HTMLParser
    from html import unescape
    class TextExtractor(HTMLParser):
        SKIP = {"script", "style", "noscript", "svg", "nav", "footer", "header", "form"}
        def __init__(self):
            super().__init__(); self.parts=[]; self.skip=0
        def handle_starttag(self, tag, attrs):
            tag=tag.lower()
            if tag in self.SKIP: self.skip += 1
            elif self.skip == 0 and tag in {"p","li","h1","h2","h3","h4","h5","h6","br","tr"}: self.parts.append("\n")
        def handle_endtag(self, tag):
            tag=tag.lower()
            if tag in self.SKIP and self.skip: self.skip -= 1
            elif self.skip == 0 and tag in {"p","li","h1","h2","h3","h4","h5","h6","br","tr"}: self.parts.append("\n")
        def handle_data(self, data):
            if self.skip == 0 and data.strip(): self.parts.append(unescape(data))
    parser=TextExtractor(); parser.feed(raw)
    return re.sub(r"\s+", " ", " ".join(parser.parts)).strip()


def _fetch_live_page(url, timeout=15, max_chars=14000):
    result={"url":url,"fetched":False,"status":"","content_type":"","text":"","error":""}
    try:
        req=urllib.request.Request(url, headers={
            "User-Agent":"Mozilla/5.0 (compatible; AHJ-Research-Assistant/5.0)",
            "Accept":"text/html,application/xhtml+xml,application/pdf;q=0.9,*/*;q=0.8",
            "Accept-Language":"en-US,en;q=0.8"})
        with urllib.request.urlopen(req, timeout=timeout) as response:
            data=response.read(2500000)
            result["status"]=str(getattr(response,"status",""))
            result["content_type"]=response.headers.get("Content-Type","")
        ctype=result["content_type"].lower()
        if "pdf" in ctype or url.lower().split("?")[0].endswith(".pdf"):
            try:
                from pypdf import PdfReader
                import io
                reader=PdfReader(io.BytesIO(data))
                chunks=[]
                for page in reader.pages[:30]:
                    txt=page.extract_text() or ""
                    if txt: chunks.append(txt)
                result["text"]=re.sub(r"\s+"," ","\n".join(chunks)).strip()[:max_chars]
            except Exception as exc:
                result["error"]=f"PDF parsing unavailable: {type(exc).__name__}"
        else:
            result["text"]=_clean_html_text(data.decode("utf-8",errors="ignore"))[:max_chars]
        result["fetched"]=bool(result["text"])
    except Exception as exc:
        result["error"]=f"{type(exc).__name__}: {exc}"
    return result


OFFICIAL_SEED_SOURCES = {
    "Oregon": [
        "https://www.oregon.gov/bcd/codes-stand/pages/adopted-codes.aspx",
        "https://www.oregon.gov/bcd/codes-stand/pages/mechanical.aspx",
        "https://www.oregon.gov/bcd/codes-stand/pages/commercial-structures.aspx",
        "https://www.oregon.gov/bcd/codes-stand/Pages/oeesc-adoption.aspx",
        "https://www.washingtoncountyor.gov/lut/building-services/building-and-development-application-services",
    ],
}


def fetch_seed_sources(state):
    """Fetch known official source landing pages live before search discovery."""
    items=[]
    retrieved_at=datetime.now().astimezone().isoformat(timespec="seconds")
    for url in OFFICIAL_SEED_SOURCES.get(state, []):
        page=_fetch_live_page(url)
        if page.get("fetched") and len(page.get("text", "")) >= 200:
            items.append({
                "title": url,
                "url": url,
                "snippet": page["text"][:700],
                "content": page["text"],
                "query": "official seed source",
                "search_engine": "direct-official-source",
                "retrieved_at": retrieved_at,
                "status": page.get("status"),
                "content_type": page.get("content_type"),
            })
    return items


def search_web_evidence(query, max_results=6):
    """Live discovery plus direct fetching. Search snippets never become evidence."""
    from urllib.parse import quote, unquote, parse_qs, urlparse
    from html import unescape
    candidates=[]
    endpoints=[
        "https://html.duckduckgo.com/html/?q="+quote(query),
        "https://www.google.com/search?q="+quote(query)+"&num=10",
        "https://www.bing.com/search?q="+quote(query)+"&count=10",
    ]
    for endpoint in endpoints:
        try:
            req=urllib.request.Request(endpoint,headers={
                "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36",
                "Accept-Language":"en-US,en;q=0.9"})
            with urllib.request.urlopen(req,timeout=12) as r:
                raw=r.read().decode("utf-8",errors="ignore")
            lowered=raw.lower()
            if len(raw)<1000 or "captcha" in lowered or "unusual traffic" in lowered: continue
            # DuckDuckGo
            pat = r"<a[^>]+class=[\"\'][^\"\']*result__a[^\"\']*[\"\'][^>]+href=[\"\']([^\"\']+)[\"\'][^>]*>(.*?)</a>"
            for m in re.finditer(pat,raw,re.I|re.S):
                u=unescape(m.group(1)).strip()
                if "uddg=" in u:
                    try: u=unquote(parse_qs(urlparse(u).query).get("uddg",[u])[0])
                    except Exception: pass
                title=re.sub(r"<[^>]+>"," ",unescape(m.group(2))).strip()
                if u.startswith(("http://","https://")): candidates.append((u,title,"DuckDuckGo"))
            # Google: collect absolute links from result anchors.
            google_pat = r"<a[^>]+href=[\"\'](https?://[^\"\']+)[\"\']"
            for m in re.finditer(google_pat, raw, re.I):
                u=unescape(m.group(1))
                host=urlparse(u).netloc.lower()
                if host and not host.endswith("google.com") and "googleusercontent.com" not in host:
                    candidates.append((u,"","Google"))
        except Exception:
            continue
    ranked=[]; seen=set()
    for u,title,engine in candidates:
        key=u.split("#",1)[0].rstrip("/").lower()
        if key in seen: continue
        seen.add(key)
        host=urlparse(u).netloc.lower(); score=0
        if host.endswith(".gov") or ".gov." in host: score+=100
        if host.endswith(".us"): score+=25
        if host.endswith(".edu"): score+=15
        if any(x in host for x in ["facebook.com","youtube.com","reddit.com","yelp.com"]): score-=50
        ranked.append((score,u,title,engine))
    ranked.sort(key=lambda x:x[0],reverse=True)
    evidence=[]; retrieved_at=datetime.now().astimezone().isoformat(timespec="seconds")
    for _,u,title,engine in ranked[:max_results*4]:
        page=_fetch_live_page(u)
        if not page["fetched"] or len(page["text"])<200: continue
        evidence.append({"title":title or u,"url":u,"snippet":page["text"][:700],"content":page["text"],"query":query,"search_engine":engine,"retrieved_at":retrieved_at,"status":page["status"],"content_type":page["content_type"]})
        if len(evidence)>=max_results: break
    return evidence


def retrieve_web_evidence(fp):
    state=fp.get("state",""); address=fp.get("address",""); pt=fp.get("project_type",""); cls=fp.get("building_classification",""); project_date=fp.get("project_date","")
    all_evidence=fetch_seed_sources(state)
    queries=[
        f'"{address}" building permit jurisdiction official government',
        f'"{address}" planning zoning land use official government',
        f'"{address}" conditional use permit CUP official government',
        f'{state} {pt} {cls} mechanical permit replacement official government',
        f'{state} mechanical permit exemption replacement HVAC official government',
        f'{state} electrical permit HVAC replacement official government',
    ]
    for q in queries:
        all_evidence.extend(search_web_evidence(q,5))
    by_url={}
    for item in all_evidence:
        url=item.get("url","")
        if not url: continue
        key=url.split("#",1)[0].rstrip("/").lower()
        if key not in by_url or len(item.get("content", ""))>len(by_url[key].get("content", "")):
            by_url[key]=item
    values=list(by_url.values())
    # Prefer actual government hosts; exclude obvious staging/mirror pages when a
    # production government page for the same subject was retrieved.
    def rank(x):
        u=x.get("url","").lower(); host=urlparse(u).netloc.lower()
        score=0
        if host.endswith(".gov"): score+=100
        if ".gov." in host: score+=80
        if host.endswith(".us"): score+=35
        if "stage." in host: score-=40
        if any(bad in host for bad in ["facebook.com","youtube.com","reddit.com","yelp.com"]): score-=100
        if "official seed source" == x.get("query"): score+=25
        return score
    values.sort(key=lambda x:(rank(x),len(x.get("content",""))),reverse=True)
    return values[:18]

def format_web_evidence(items):
    if not items: return "NO LIVE WEB EVIDENCE WAS RETRIEVED. Current claims must be UNKNOWN/VERIFY."
    blocks=[]
    for i,x in enumerate(items,1):
        blocks.append(
            f"[LIVE SOURCE {i}]\nTitle: {x.get('title','')}\nURL: {x.get('url','')}\nRetrieved: {x.get('retrieved_at','')}\nSearch discovery query: {x.get('query','')}\nParsed page content:\n{x.get('content','')[:14000]}"
        )
    return "\n\n".join(blocks)


# -----------------------------
# Research prompts
# -----------------------------
def discovery_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

EXTERNAL WEB EVIDENCE:
{format_web_evidence(st.session_state.get("web_evidence", []))}

RESEARCH PASS 1 — JURISDICTION AND OFFICIAL SOURCE MAP

Find the authoritative government sources needed to research this project.

Determine:
1. State-level building/code authority.
2. Most likely local permitting authority.
3. Building department / permit portal.
4. Fire/life-safety authority when relevant.
5. Planning/zoning authority when relevant.
6. Any official source needed to verify jurisdiction from the address.

Then determine which code families are potentially relevant.

IMPORTANT:
Use the EXTERNAL WEB EVIDENCE supplied to you. Do not claim to have independently searched the web.
- Prefer official government sources.
- Do not assume the local jurisdiction from the city name alone.
- If the address does not establish the AHJ, say so.
- Give URLs and explain what each source is for.

Return concise findings.
"""


def code_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

EXTERNAL WEB EVIDENCE:
{format_web_evidence(st.session_state.get("web_evidence", []))}

RESEARCH PASS 2 — CURRENT APPLICABLE CODE CYCLES

This is the highest-priority pass.

Search the official state code/building agency first.

Determine, for this project and project date:

- building/structural code
- existing-building code if relevant
- mechanical code
- electrical code
- plumbing code if relevant
- energy code
- fire/life-safety code if separately adopted
- state/local amendments
- effective dates
- mandatory/adoption dates
- phase-in or transition provisions

For every code family you identify, report:

Code family | Adopted edition | Effective/mandatory date | Official source | Why it applies

CRITICAL:
- Do NOT use model memory for the code edition.
- Do NOT assume the state uses the latest IBC/IMC/IFC/etc.
- Do NOT use a neighboring state's code.
- Do NOT call a code current unless an official source supports it.
- If the official source cannot establish the edition, say UNKNOWN/VERIFY.
- If the project date falls in a transition period, explain that.

For an existing-building replacement, specifically investigate whether
replacement/alteration provisions differ from new construction.

Return concise findings with source URLs.
"""


def scope_prompt(fp):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

EXTERNAL WEB EVIDENCE:
{format_web_evidence(st.session_state.get("web_evidence", []))}

RESEARCH PASS 3 — PROJECT-SPECIFIC PERMITS AND CODE TRIGGERS

Research only the requirements actually relevant to this scope.

For each potentially relevant issue determine:

Issue | Preliminary result | Trigger/fact | Authority | Source

Potential categories include only when supported by scope:
- building permit
- mechanical permit
- electrical permit
- plumbing permit
- fire permit/review
- planning/zoning review
- CUP or land-use conditions
- accessibility
- egress
- equipment replacement
- energy efficiency
- seismic/anchorage
- structural alterations
- site work
- rooftop/roof penetrations
- hazardous materials/refrigerant
- other scope-specific issues

Do not assume every category applies.

If a permit or review cannot be established from the available evidence,
say UNKNOWN/VERIFY and formulate the exact AHJ question.

Search official local/state sources.
Return concise findings with source URLs.
"""


def conflict_prompt(fp, prior_notes):
    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

EXTERNAL WEB EVIDENCE:
{format_web_evidence(st.session_state.get("web_evidence", []))}

RESEARCH PASS 4 — ERROR CHECK / CONFLICT CHECK

Below are preliminary research notes from other passes.

{prior_notes}

Treat these as untrusted research notes.

Search current authoritative sources and look specifically for:
- stale code editions
- wrong effective dates
- residential/commercial mix-ups
- wrong permitting authority
- wrong fire authority
- unsupported permit claims
- incorrect assumptions about existing-building work
- unsupported zoning/CUP statements
- sources that contradict each other

Return only:
1. Conflict or possible error
2. Correct/current evidence
3. Source
4. What the volunteer should do if unresolved

Do not simply agree with the notes.
"""


def synthesis_prompt(fp, research_passes, source_index):
    source_index = list(source_index) + st.session_state.get("web_evidence", [])
    joined = "\n\n".join(
        f"===== RESEARCH PASS {i + 1} =====\n{item['text']}"
        for i, item in enumerate(research_passes)
    )

    source_lines = []
    for i, source in enumerate(source_index, 1):
        source_lines.append(
            f"[S{i}] {source.get('title', 'Source')} — {source.get('url', '')}"
        )

    sources_text = "\n".join(source_lines)

    return f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

RESEARCH NOTES:
{joined}

SOURCES ACTUALLY RETRIEVED:
{sources_text}

Now synthesize the final volunteer-facing report.

IMPORTANT:
- The research notes are evidence summaries, not automatic truth.
- Prefer the strongest authoritative source.
- If notes conflict, resolve them using the source hierarchy and current date.
- If the evidence is insufficient, say UNKNOWN or VERIFY.
- Do not invent facts to make the report complete.
- Do not state a code edition without source support.
- Do not call a claim VERIFIED unless the retrieved evidence actually supports it.
- Every important code/permit conclusion should cite one or more source IDs
  such as [S3].
- Use only source IDs that appear in SOURCES ACTUALLY RETRIEVED.

FORMAT EXACTLY:

# PRELIMINARY ANSWER

Start with the practical answer.

| Issue | Result | Confidence |
|---|---|---|

Only include issues relevant to this project.

# CURRENT CODE PATH

Give the applicable code families and editions discovered for this project.
Include effective/mandatory dates where verified.
Cite source IDs.

# WHAT WE KNOW

Maximum 8 bullets.
Cite important claims.

# WHAT COULD CHANGE THE ANSWER

Maximum 5 bullets.

# QUESTIONS FOR THE AHJ

Maximum 5 precise questions.
Only ask questions that could change the result.

# VOLUNTEER ACTION LIST

Maximum 6 steps, in order.

# SOURCES

List only sources actually retrieved.
Use:
- [S1] Title — URL — what it supports

# RESEARCH LIMITATIONS

One short paragraph explaining what could not be independently established.

Do not repeat disclaimers throughout the report.
"""


def clean_report(text):
    if not text:
        return text

    text = re.sub(r"\n{3,}", "\n\n", text)
    text = text.replace("🟢 CONFIRMED", "🟢 VERIFIED")
    return text.strip()


def source_diagnostics(report, sources):
    urls = []

    for source in sources:
        url = source.get("url")
        if url and url not in urls:
            urls.append(url)

    # Also catch URLs manually written into the final answer.
    for url in extract_urls(report):
        if url not in urls:
            urls.append(url)

    checks = []
    for url in urls[:30]:
        checks.append(check_url(url))

    return checks


def build_docx(fp, report, verified_facts, diagnostics, sources):
    doc = Document()

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)

    doc.add_heading("AHJ Research Report", 0)
    doc.add_paragraph(
        f"Project: {fp.get('project_type', '')}\n"
        f"State: {fp.get('state', '')}\n"
        f"Address: {fp.get('address', '')}\n"
        f"Project date: {fp.get('project_date', '')}\n"
        f"Generated: {datetime.now().strftime('%B %d, %Y')}"
    )

    doc.add_paragraph(
        "DRAFT RESEARCH AID — Verify applicable requirements with the AHJ "
        "before relying on this report."
    )

    doc.add_heading("Project Scope", level=1)
    doc.add_paragraph(compact_fingerprint(fp))

    doc.add_heading("Research Report", level=1)

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
            doc.add_paragraph(line)
        else:
            doc.add_paragraph(line)

    if sources:
        doc.add_heading("Live Sources Retrieved and Parsed", level=1)
        for i, source in enumerate(sources, 1):
            doc.add_paragraph(
                f"[S{i}] {source.get('title', 'Source')}\n"
                f"{source.get('url', '')}",
                style="List Bullet",
            )

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


def save_session(fp, report, notes, sources):
    return json.dumps(
        {
            "saved": datetime.now().isoformat(),
            "project": fp,
            "report": report,
            "verified_facts": notes,
            "sources": sources,
        },
        indent=2,
    )


# ============================================================
# UI
# ============================================================

st.title("🏛️ AHJ Research Assistant")
st.caption(
    "Search-grounded research. Official sources first. "
    "State-agnostic code discovery. Human verification for final decisions."
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

    st.subheader("Search grounding")
    st.info(
        "Gemini is instructed to search the current web for code editions, "
        "effective dates, AHJs, permits, and local requirements. Code cycles "
        "are no longer hard-coded into the application."
    )

# -----------------------------
# Project inputs
# -----------------------------
st.header("1. Project")

col1, col2 = st.columns(2)

with col1:
    state = st.selectbox(
        "State / jurisdiction",
        STATE_OPTIONS,
        index=STATE_OPTIONS.index("Oregon"),
    )

    address = st.text_input(
        "Project address",
        value="24340 NW Meek Rd, 97124",
    )

    project_date = st.date_input(
        "Project / permit date",
        value=date.today(),
        help=(
            "Use the date that matters for code applicability. "
            "If unknown, use the expected permit/application date."
        ),
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

    building_status = st.selectbox(
        "Building status",
        [
            "Existing building",
            "Existing building — alteration/remodel",
            "New construction",
            "Addition",
            "Unknown",
        ],
        index=0,
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

scope = {
    "building_status": building_status,
}

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
    state=state,
    project_date=project_date,
    project_type=project_type,
    classification=classification,
    address=address,
    details=details,
    scope_answers=scope,
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
        "The research engine searches the current web instead of relying on "
        "hard-coded state code cycles. Deep Research performs separate "
        "jurisdiction, code, scope, and conflict passes."
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
        with st.spinner(
            "Retrieving live web pages, parsing official sources, and analyzing project-specific issues..."
        ):
            # Retrieve evidence BEFORE building Gemini prompts. The prior version
            # defined retrieve_web_evidence() but never actually called it, which
            # guaranteed 0 sources and left Gemini with no current evidence.
            web_evidence = retrieve_web_evidence(fp)
            st.session_state["web_evidence"] = web_evidence
            passes = []

            # This application is research-first: if live retrieval fails, do
            # not let Gemini manufacture a report from its training memory.
            if not web_evidence:
                result = {
                    "text": (
                        "# RESEARCH COULD NOT BE COMPLETED\n\n"
                        "No live web pages could be retrieved from the external "
                        "research layer. Gemini was intentionally NOT called, "
                        "because this application requires current evidence before "
                        "making code, permit, jurisdiction, or zoning conclusions.\n\n"
                        "Check Streamlit Cloud network access to the public search "
                        "endpoints and official government sites, then run the "
                        "research again."
                    ),
                    "sources": [],
                    "model": "",
                    "error": True,
                }
                passes = [result]
            elif depth == "Quick Answer":
                prompt = f"""
{base_system_rules()}

PROJECT:
{compact_fingerprint(fp)}

EXTERNAL WEB EVIDENCE:
{format_web_evidence(web_evidence)}

Perform a concise but current research pass.

Use only the EXTERNAL WEB EVIDENCE supplied below. Do not perform web search from Gemini.

First identify the authoritative state code source and the likely local AHJ.
Then determine the current applicable code path for the stated project date.
Then determine the few permit/code/planning issues that matter to this scope.

Return:

# PRELIMINARY ANSWER

| Issue | Result | Confidence |
|---|---|---|

# CURRENT CODE PATH

List applicable code families, editions, and dates with source citations.

# WHAT WE KNOW

Use 8-12 bullets when the evidence supports them. Distinguish VERIFIED, LIKELY, CONDITIONAL, UNKNOWN, and VERIFY.

# WHAT COULD CHANGE THE ANSWER

Use up to 8 bullets.

# QUESTIONS FOR THE AHJ

Use up to 7 specific questions.

# SOURCES

List only the supplied LIVE SOURCE numbers actually relied upon, e.g. [LIVE SOURCE 2]. Do not invent URLs or a separate Gemini source list.

# RESEARCH LIMITATIONS

Short paragraph.

Never rely on stale model memory when the current web can answer it.
Do NOT state a permit requirement, code edition, AHJ, exemption, or technical
code section as VERIFIED unless a supplied source supports that exact claim.
If the supplied evidence is insufficient, mark the issue UNKNOWN or VERIFY.
Preserve useful technical detail from the project scope; do not collapse the
report into generic advice.
"""
                result = call_gemini(prompt)
                passes = [result]

            else:
                evidence_block = format_web_evidence(web_evidence)
                prompts = [
                    discovery_prompt(fp) + "\n\nEXTERNAL WEB EVIDENCE:\n" + evidence_block,
                    code_prompt(fp) + "\n\nEXTERNAL WEB EVIDENCE:\n" + evidence_block,
                    scope_prompt(fp) + "\n\nEXTERNAL WEB EVIDENCE:\n" + evidence_block,
                ]

                progress = st.progress(0)

                for i, prompt in enumerate(prompts):
                    result = call_gemini(prompt)
                    passes.append(result)
                    progress.progress((i + 1) / (len(prompts) + 1))

                prior_notes = "\n\n".join(
                    f"PASS {i + 1}:\n{item['text']}"
                    for i, item in enumerate(passes)
                )

                conflict = call_gemini(
                    conflict_prompt(fp, prior_notes),
                )
                passes.append(conflict)
                progress.progress(4 / 5)

                # Source truth comes ONLY from the live retrieval layer.
                # Model-generated URLs are never promoted into the source index.
                source_index = list(st.session_state.get("web_evidence", []))

                synthesis = call_gemini(
                    synthesis_prompt(fp, passes, source_index)
                    + "\n\nEXTERNAL WEB EVIDENCE:\n"
                    + format_web_evidence(web_evidence),
                )

                # The synthesis call intentionally does not search again.
                # This keeps the final answer from introducing an uncited new
                # code edition after the grounded research passes.
                final_result = synthesis

                progress.progress(1.0)

                passes.append(final_result)

                result = final_result

        st.session_state["research_log"] = passes

        # Build final source index from the external retrieval layer plus
        # any explicitly identified sources.
        all_sources = list(st.session_state.get("web_evidence", []))

        for item in passes:
            for source in item.get("sources", []):
                url = source.get("url")
                if not url:
                    continue

                if url not in {s.get("url") for s in all_sources}:
                    all_sources.append(source)

        st.session_state["research_sources"] = all_sources
        st.session_state["report"] = clean_report(result["text"])

        with st.spinner("Checking source URLs for basic reachability..."):
            st.session_state["source_checks"] = source_diagnostics(
                st.session_state["report"],
                st.session_state["research_sources"],
            )

        st.success(
            f"Research complete. "
            f"{len(st.session_state['research_sources'])} source(s) were retrieved."
        )

# -----------------------------
# Results
# -----------------------------
if st.session_state["report"]:
    st.header("4. Preliminary Result")

    st.markdown(st.session_state["report"])

    if st.session_state["research_sources"]:
        with st.expander("Sources actually retrieved by Gemini"):
            for i, source in enumerate(st.session_state["research_sources"], 1):
                st.markdown(
                    f"**[S{i}] {source.get('title', 'Source')}**  \n"
                    f"{source.get('url', '')}"
                )

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
        "Record facts personally confirmed with an AHJ or official record. "
        "These facts can be included in the Word report."
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
