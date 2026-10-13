"""Scope extraction with Google Gemini (google-genai SDK, native structured output).

Role of the model, and its limits
---------------------------------
The model only *reads and classifies*: it proposes scope items, a WBS position, a duration
basis and candidate logic, each with a citation (segment id + verbatim quote) to the source
document. It never produces dates. Everything it returns is:

1. parsed into the ``ProjectScope`` Pydantic schema (``response_schema=ProjectScope``),
2. normalised deterministically (unique refs, milestone durations, no SF / leads),
3. citation-checked against the document text (``verify_citations``), and
4. put in front of a planner in the Review Gate, who must approve it before any CPM logic runs.

Security: the API key is passed in by the caller (Streamlit secrets or a masked sidebar field);
it is never logged, stored or written to any output.
"""
from __future__ import annotations

import json
import random
import re
import time
from dataclasses import dataclass, field
from typing import Literal, Optional

import pandas as pd
from pydantic import BaseModel, Field, ValidationError

from .doc_reader import DocumentText, normalise_for_match
from .scope_builder import REGISTER_COLUMNS, format_predecessors

DEFAULT_MODEL = "gemini-2.5-flash"
MAX_OUTPUT_TOKENS = 65_536
THINKING_BUDGET = 4_096          # small reasoning budget: better structure, bounded cost and latency
RETRIES = 3


# --------------------------------------------------------------------------- #
# Response schema (passed to Gemini as response_schema=ProjectScope)
# No default values: every field is required (nullable where it may be unknown), which keeps the
# schema inside what Gemini structured output supports.
# --------------------------------------------------------------------------- #
class Dependency(BaseModel):
    predecessor_ref: str = Field(description="ref of the item that must come first (e.g. A020)")
    relationship: Literal["FS", "SS", "FF"] = Field(description="FS unless the document clearly implies overlap")
    lag_days: int = Field(description="working days of lag, 0 if none; never negative")


class ScopeItem(BaseModel):
    ref: str = Field(description="unique reference: A010, A020, ... in steps of 10, in document order")
    wbs: str = Field(description="numeric outline code: 1, 1.1, 1.1.1 (max 4 levels)")
    name: str = Field(description="verb-noun activity name, max 80 characters, e.g. 'Install pump skid'")
    item_type: Literal["Summary", "Task", "Milestone"]
    duration_days: Optional[int] = Field(description="working days; null for Summary; 0 for Milestone")
    duration_basis: Literal["stated", "estimated", "not_applicable"] = Field(
        description="stated = the document gives it; estimated = your planning estimate; not_applicable = "
                    "Summary or Milestone")
    predecessors: list[Dependency] = Field(description="empty for Summary rows and for the first milestone")
    source_ref: str = Field(description="segment id in square brackets in the document, without brackets, "
                                        "e.g. P14 or T2.R5 or PG3.4")
    source_quote: str = Field(description="exact words copied from that segment (5 to 25 words) that justify "
                                          "this item")
    confidence: float = Field(description="0.0 to 1.0: how directly the document supports the item and its "
                                         "duration")
    assumption: Optional[str] = Field(description="planning assumption behind an estimated duration or "
                                                  "inferred logic; null if none")


class ProjectScope(BaseModel):
    project_name: str = Field(description="project title as written in the document")
    project_start_date: Optional[str] = Field(description="YYYY-MM-DD if the document states a start, "
                                                          "commencement, award or NTP date; else null")
    work_week_days: Optional[int] = Field(description="5, 6 or 7 if the document states the working week; "
                                                      "else null")
    items: list[ScopeItem]
    exclusions: list[str] = Field(description="scope the document explicitly excludes")
    open_questions: list[str] = Field(description="gaps or ambiguities a planner must resolve")


# --------------------------------------------------------------------------- #
# Prompt
# --------------------------------------------------------------------------- #
DETAIL_LEVELS = {
    "L2": "Level 2 (management summary): phases and major deliverables, about 20-80 items.",
    "L3": "Level 3 (contractor control): work packages by discipline/area, about 60-300 items.",
}

SYSTEM_INSTRUCTION = """You are a senior planning engineer (PMI-SP, AACE PSP) building a scope register
for a capital project (oil & gas, infrastructure, power, engineering services) that will become a CPM
baseline in Primavera P6 and MS Project.

Rules:
1. Use ONLY the scope in the document. Do not invent deliverables. If something is implied but not
   written, add it to open_questions instead of the items list.
2. Every item MUST cite source_ref = the segment id shown in square brackets at the start of a line,
   and source_quote = words copied exactly from that segment (5-25 words). Never paraphrase the quote.
3. Structure: Summary rows for phases/areas (WBS 1, 2, 3 ...), Task and Milestone rows beneath them
   (1.1, 1.2 ...). Order items by WBS. A Summary must have at least one child. Max 4 WBS levels.
   Typical capital-project phases: mobilisation/start-up, engineering, procurement, construction/
   fabrication, installation, pre-commissioning/commissioning, handover/close-out (use the ones the
   document supports).
4. Start with a start milestone (e.g. Contract Award / Notice to Proceed / Kick-off) as the first
   non-summary item when the document mentions award, NTP, commencement or kick-off.
5. Milestones: 0 days, duration_basis not_applicable. Summary: duration null, not_applicable.
   Tasks: integer working days > 0. Use durations stated in the document (convert weeks x5, months x22)
   with basis 'stated'; otherwise give a realistic planning estimate with basis 'estimated' and state
   the assumption.
6. Logic: refer only to refs in your own list. Prefer FS with 0 lag. Use SS with a positive lag only
   for genuine overlap. FF only where finishes are tied. Never SF. Never negative lag (leads).
   No predecessors on Summary rows. Every Task and Milestone except the first start milestone needs at
   least one predecessor. Avoid circular logic.
7. Names: verb + noun, specific, max 80 characters, no WBS numbers in the name.
8. confidence: 0.9+ when the item and its duration are written explicitly; 0.6-0.8 when the item is
   written but the duration is estimated; below 0.6 when the item is inferred from context.
9. project_start_date only if a date is written in the document (YYYY-MM-DD)."""


def build_prompt(doc: DocumentText, detail: str, text: str) -> str:
    return (f"Target detail: {DETAIL_LEVELS.get(detail, DETAIL_LEVELS['L2'])}\n"
            f"Document: {doc.filename} ({doc.file_type}, {doc.word_count:,} words, "
            f"{len(doc.segments)} segments).\n"
            "Each line starts with its segment id in square brackets.\n\n"
            "=== DOCUMENT START ===\n"
            f"{text}\n"
            "=== DOCUMENT END ===\n\n"
            "Return the scope register as JSON matching the schema.")


# --------------------------------------------------------------------------- #
# Redaction (client confidentiality)
# --------------------------------------------------------------------------- #
def redact(text: str, terms: list) -> tuple:
    """Replace each confidential term (client name, project code, person) with a neutral token."""
    count = 0
    for i, term in enumerate(t for t in (terms or []) if t and len(t.strip()) >= 2):
        token = f"[REDACTED-{i + 1}]"
        text, n = re.subn(re.escape(term.strip()), token, text, flags=re.I)
        count += n
    return text, count


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #
class ExtractionError(RuntimeError):
    """User-facing extraction failure."""


@dataclass
class Extraction:
    scope: ProjectScope
    model: str
    detail: str
    seconds: float
    input_tokens: Optional[int] = None
    output_tokens: Optional[int] = None
    redactions: int = 0
    notes: list = field(default_factory=list)


def _client(api_key: str):
    try:
        from google import genai
    except ImportError as exc:  # pragma: no cover
        raise ExtractionError("The google-genai package is not installed (see requirements.txt).") from exc
    return genai.Client(api_key=api_key)


def extract_scope(doc: DocumentText, api_key: str, model: str = DEFAULT_MODEL, detail: str = "L2",
                  redact_terms: Optional[list] = None, client=None) -> Extraction:
    """Call Gemini once with structured output; retry transient errors. Raises ExtractionError."""
    if not api_key and client is None:
        raise ExtractionError("No Gemini API key. Add GEMINI_API_KEY to Streamlit secrets or enter it in the sidebar.")
    from google.genai import errors, types

    text, n_red = redact(doc.as_prompt_text(), redact_terms or [])
    prompt = build_prompt(doc, detail, text)
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_INSTRUCTION,
        response_mime_type="application/json",
        response_schema=ProjectScope,
        temperature=0.0,
        seed=7,
        max_output_tokens=MAX_OUTPUT_TOKENS,
        thinking_config=types.ThinkingConfig(thinking_budget=THINKING_BUDGET),
    )
    client = client or _client(api_key)

    t0 = time.time()
    last_exc = None
    resp = None
    for attempt in range(1, RETRIES + 1):
        try:
            resp = client.models.generate_content(model=model, contents=prompt, config=config)
            break
        except errors.ClientError as exc:                 # 4xx
            code = getattr(exc, "code", None)
            if code == 429 and attempt < RETRIES:          # rate limit: back off and retry
                last_exc = exc
                time.sleep(min(30, 4 * 2 ** (attempt - 1)) + random.random())
                continue
            raise ExtractionError(_client_error_message(exc)) from exc
        except errors.ServerError as exc:                 # 5xx: transient
            last_exc = exc
            if attempt < RETRIES:
                time.sleep(min(30, 4 * 2 ** (attempt - 1)) + random.random())
                continue
            raise ExtractionError(f"Gemini service unavailable after {RETRIES} attempts ({exc}). "
                                  "Try again in a few minutes.") from exc
        except Exception as exc:  # noqa: BLE001 - network, DNS, TLS
            last_exc = exc
            if attempt < RETRIES:
                time.sleep(2 * attempt)
                continue
            raise ExtractionError(f"Could not reach the Gemini API: {exc}") from exc
    if resp is None:  # pragma: no cover
        raise ExtractionError(f"Gemini call failed: {last_exc}")

    finish = _finish_reason(resp)
    if finish == "MAX_TOKENS":
        raise ExtractionError("The scope is too large for one extraction (the model's output limit was reached). "
                              "Choose Level 2 detail, or split the document into parts and run each one.")
    if finish in ("SAFETY", "RECITATION", "PROHIBITED_CONTENT", "BLOCKLIST", "SPII"):
        raise ExtractionError(f"Gemini declined to return the extraction (finish reason {finish}).")

    scope = getattr(resp, "parsed", None)
    if not isinstance(scope, ProjectScope):
        raw = getattr(resp, "text", None) or ""
        try:
            scope = ProjectScope.model_validate_json(raw)
        except (ValidationError, ValueError) as exc:
            raise ExtractionError("Gemini returned output that does not match the scope schema. "
                                  "Run the extraction again.") from exc
    if not scope.items:
        raise ExtractionError("No scope items were found in the document.")

    usage = getattr(resp, "usage_metadata", None)
    ext = Extraction(scope=scope, model=model, detail=detail, seconds=round(time.time() - t0, 1),
                     input_tokens=getattr(usage, "prompt_token_count", None) if usage else None,
                     output_tokens=getattr(usage, "candidates_token_count", None) if usage else None,
                     redactions=n_red)
    ext.notes.extend(normalise_scope(scope))
    return ext


def _finish_reason(resp) -> Optional[str]:
    try:
        fr = resp.candidates[0].finish_reason
    except (AttributeError, IndexError, TypeError):
        return None
    if fr is None:
        return None
    return getattr(fr, "name", str(fr)).split(".")[-1]


def _client_error_message(exc) -> str:
    code = getattr(exc, "code", None)
    if code in (401, 403):
        return "The Gemini API key was rejected (401/403). Check GEMINI_API_KEY."
    if code == 404:
        return "The Gemini model was not found (404). Check GEMINI_MODEL in secrets."
    if code == 429:
        return "Gemini rate limit or quota exceeded (429). Wait a minute, or check the quota on the API key."
    if code == 400:
        return f"Gemini rejected the request (400): {getattr(exc, 'message', exc)}"
    return f"Gemini request failed ({code}): {getattr(exc, 'message', exc)}"


# --------------------------------------------------------------------------- #
# Deterministic normalisation of the model output
# --------------------------------------------------------------------------- #
def normalise_scope(scope: ProjectScope) -> list:
    """Make the register safe to review: unique refs, clean types/durations, no SF/leads. Returns notes."""
    notes = []
    seen = set()
    rename = {}
    for i, it in enumerate(scope.items):
        ref = re.sub(r"[^A-Za-z0-9._\-]", "", (it.ref or "").upper()) or f"A{(i + 1) * 10:03d}"
        base, k = ref, 1
        while ref in seen:
            k += 1
            ref = f"{base}-{k}"
        if ref != (it.ref or "").upper():
            rename[(it.ref or "").upper()] = ref
            if base != ref:
                notes.append(f"Duplicate ref {base} renamed {ref}.")
        seen.add(ref)
        it.ref = ref
        it.wbs = (it.wbs or "").strip().rstrip(".")
        it.name = re.sub(r"\s+", " ", it.name or "").strip()[:120]
        it.confidence = max(0.0, min(1.0, float(it.confidence or 0)))
        if it.item_type == "Milestone":
            it.duration_days, it.duration_basis = 0, "not_applicable"
        elif it.item_type == "Summary":
            it.duration_days, it.duration_basis = None, "not_applicable"
            if it.predecessors:
                notes.append(f"{ref}: logic removed from summary row.")
                it.predecessors = []
    for it in scope.items:
        for d in it.predecessors:
            d.predecessor_ref = rename.get((d.predecessor_ref or "").upper(), (d.predecessor_ref or "").upper())
            if d.lag_days < 0:
                notes.append(f"{it.ref}: lead of {d.lag_days}d from {d.predecessor_ref} set to 0.")
                d.lag_days = 0
    return notes


# --------------------------------------------------------------------------- #
# Citation verification (traceability)
# --------------------------------------------------------------------------- #
def verify_citation(quote: str, source_ref: str, doc: DocumentText) -> tuple:
    """-> (status, source_ref). Verified = quote found in the cited segment; Relocated = found in another
    segment (ref corrected); Unverified = not found anywhere (possible hallucination)."""
    q = normalise_for_match(quote)
    if len(q) < 8:
        return "Unverified", source_ref
    segs = doc.by_id()
    sid = (source_ref or "").strip().strip("[]")
    if sid in segs and q in normalise_for_match(segs[sid].text):
        return "Verified", sid
    for s in doc.segments:
        if q in normalise_for_match(s.text):
            return "Relocated", s.sid
    # tolerate a quote that spans two adjacent segments
    ids = [s.sid for s in doc.segments]
    if sid in segs:
        i = ids.index(sid)
        window = " ".join(normalise_for_match(doc.segments[j].text) for j in range(max(0, i - 1), min(len(ids), i + 2)))
        if q in window:
            return "Verified", sid
    return "Unverified", sid


def scope_to_register(scope: ProjectScope, doc: DocumentText) -> pd.DataFrame:
    """Model output -> Review Gate register (one row per item)."""
    rows = []
    for it in scope.items:
        status, sid = verify_citation(it.source_quote, it.source_ref, doc)
        deps = [(d.predecessor_ref, d.relationship, d.lag_days) for d in it.predecessors]
        rows.append({
            "Include": True,
            "Ref": it.ref,
            "WBS": it.wbs,
            "Activity": it.name,
            "Type": it.item_type,
            "Duration (days)": None if it.item_type == "Summary" else it.duration_days,
            "Predecessors": format_predecessors(deps),
            "Basis": {"stated": "Stated", "estimated": "Estimated"}.get(it.duration_basis, "-"),
            "Confidence": round(it.confidence, 2),
            "Citation": status,
            "Source": sid,
            "Evidence": it.source_quote,
            "Assumption": it.assumption or "",
            "Planner note": "",
        })
    df = pd.DataFrame(rows, columns=REGISTER_COLUMNS)
    return df


def scope_summary(scope: ProjectScope) -> dict:
    return json.loads(scope.model_dump_json(include={"project_name", "project_start_date", "work_week_days",
                                                      "exclusions", "open_questions"}))
