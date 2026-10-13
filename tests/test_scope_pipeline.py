"""Regression tests for document scope ingestion (no network, no API key, no client data).

Covers: document reading, Gemini structured-output handling (mocked client), citation
verification, Review Gate validation, the deterministic register -> engine bridge
(zero-day variance), firm-policy close-out (no dangling ends) and DCMA-14.

Run:  pytest -q
"""
import datetime as dt
import io
import os
import sys
from types import SimpleNamespace

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scheduler import cpm, dcma, exporters  # noqa: E402
from scheduler import extractor as ex  # noqa: E402
from scheduler import scope_builder as sb  # noqa: E402
from scheduler.doc_reader import DocumentReadError, read_document  # noqa: E402

D0 = dt.date(2026, 11, 2)  # a Monday


# --------------------------------------------------------------------------- #
# Fixtures
# --------------------------------------------------------------------------- #
def make_docx() -> bytes:
    docx = pytest.importorskip("docx")
    d = docx.Document()
    d.add_heading("Statement of Work - Water Tank", 0)
    d.add_paragraph("Work shall commence on 2 November 2026 following Contract Award.")
    d.add_paragraph("Engineering: prepare the design basis within 2 weeks of award.", style="List Bullet")
    d.add_paragraph("Construction: civil foundations (4 weeks), followed by hydrostatic testing (1 week).")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Milestone", "Target"
    t.cell(1, 0).text, t.cell(1, 1).text = "Handover to Client", "Week 12"
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def register(rows):
    """rows: (ref, wbs, name, type, dur, preds)"""
    return pd.DataFrame([{"Include": True, "Ref": r, "WBS": w, "Activity": n, "Type": t, "Duration (days)": d,
                          "Predecessors": p, "Basis": "Stated", "Confidence": 0.9, "Citation": "Verified",
                          "Source": "P1", "Evidence": "quote", "Assumption": "", "Planner note": ""}
                         for r, w, n, t, d, p in rows], columns=sb.REGISTER_COLUMNS)


BASE = [
    ("A010", "1", "Start-up", "Summary", None, ""),
    ("A020", "1.1", "Contract award", "Milestone", 0, ""),
    ("A030", "1.2", "Mobilise team", "Task", 10, "A020"),
    ("A040", "1.3", "Kick-off meeting", "Task", 1, "A020"),
    ("A050", "1.4", "Kick-off minutes issued", "Milestone", 0, "A040"),
    ("A100", "2", "Engineering", "Summary", None, ""),
    ("A110", "2.1", "Design basis", "Task", 10, "A030"),
    ("A120", "2.2", "Detailed design", "Task", 30, "A110"),
    ("A200", "3", "Construction", "Summary", None, ""),
    ("A210", "3.1", "Civil works", "Task", 20, "A120"),
    ("A220", "3.2", "Hydrotest", "Task", 5, "A210"),
    ("A230", "3.3", "Coating touch-up", "Task", 3, "A220 FF"),
    ("A300", "4", "Handover", "Milestone", 0, "A200, A050"),
]


def build(rows, **kw):
    return sb.build_schedule(register(rows), sb.BuildSettings("Test Project", D0, **kw))


# --------------------------------------------------------------------------- #
# Document reader
# --------------------------------------------------------------------------- #
def test_docx_segments_headings_and_tables():
    doc = read_document(make_docx(), "sow.docx")
    ids = [s.sid for s in doc.segments]
    assert ids[0] == "P1" and doc.segments[0].kind == "heading"
    assert "T1.R2" in ids and doc.by_id()["T1.R2"].text == "Handover to Client | Week 12"
    assert any(s.text.startswith("- Engineering") for s in doc.segments)   # bullet kept as a list item
    assert "[P2] Work shall commence" in doc.as_prompt_text()


def test_pdf_text_layer():
    pytest.importorskip("reportlab")
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.drawString(72, 760, "1. Engineering: design basis within 2 weeks of award.")
    c.drawString(72, 740, "2. Construction: civil foundations (4 weeks).")
    c.showPage()
    c.save()
    doc = read_document(buf.getvalue(), "sow.pdf")
    assert doc.pages == 1 and doc.segments[0].sid == "PG1.1"
    assert "civil foundations" in doc.as_prompt_text()


def test_text_file_and_rejections():
    doc = read_document(b"# Scope\n\nDesign basis within 2 weeks of award.\n\nCivil works 4 weeks.\n", "s.md")
    assert [s.sid for s in doc.segments] == ["L1", "L2", "L3"] and doc.segments[0].kind == "heading"
    for data, name in [(b"", "a.docx"), (b"x", "a.doc"), (b"x", "a.xlsx"), (b"not a pdf", "a.pdf"),
                       (b"tiny", "a.txt")]:
        with pytest.raises(DocumentReadError):
            read_document(data, name)


# --------------------------------------------------------------------------- #
# Extraction (mocked Gemini client)
# --------------------------------------------------------------------------- #
def _scope(**over):
    item = lambda ref, wbs, name, t, dur, preds, src, quote, conf=0.9: ex.ScopeItem(  # noqa: E731
        ref=ref, wbs=wbs, name=name, item_type=t, duration_days=dur,
        duration_basis="stated" if t == "Task" else "not_applicable",
        predecessors=[ex.Dependency(predecessor_ref=p, relationship=r, lag_days=lag) for p, r, lag in preds],
        source_ref=src, source_quote=quote, confidence=conf, assumption=None)
    data = dict(project_name="Water Tank", project_start_date="2026-11-02", work_week_days=None,
                exclusions=[], open_questions=[], items=[
                    item("A010", "1", "Start", "Summary", 5, [("X", "FS", 0)], "P2", "following Contract Award"),
                    item("A020", "1.1", "Contract award", "Milestone", 3, [], "P2",
                         "commence on 2 November 2026 following Contract Award"),
                    item("A030", "1.2", "Prepare design basis", "Task", 10, [("A020", "FS", -3)], "P2",
                         "prepare the design basis within 2 weeks of award"),          # wrong segment -> relocated
                    item("A030", "1.3", "Build foundations", "Task", 20, [("A030", "FS", 0)], "P4",
                         "civil foundations (4 weeks)"),
                    item("A040", "1.4", "Invented item", "Task", 5, [], "P9", "this text is not in the document",
                         0.4),
                ])
    data.update(over)
    return ex.ProjectScope(**data)


class _Models:
    def __init__(self, scope, finish="STOP"):
        self.scope, self.finish, self.calls = scope, finish, []

    def generate_content(self, model, contents, config):
        self.calls.append((model, contents, config))
        return SimpleNamespace(parsed=self.scope, text=self.scope.model_dump_json(),
                               candidates=[SimpleNamespace(finish_reason=SimpleNamespace(name=self.finish))],
                               usage_metadata=SimpleNamespace(prompt_token_count=100, candidates_token_count=50))


def test_extract_structured_output_redaction_and_normalisation():
    doc = read_document(make_docx(), "sow.docx")
    models = _Models(_scope())
    e = ex.extract_scope(doc, "k", detail="L2", redact_terms=["Water Tank"], client=SimpleNamespace(models=models))
    model, contents, config = models.calls[0]
    assert model == "gemini-3.8-flash" and e.model == "gemini-3.8-flash"      # auto: newest free Flash first
    assert config.response_schema is ex.ProjectScope and config.response_mime_type == "application/json"
    assert config.temperature is None and config.thinking_config.thinking_level is not None   # Gemini 3 rules
    assert "Water Tank" not in contents and "[REDACTED-1]" in contents and e.redactions == 1
    items = {it.ref: it for it in e.scope.items}
    assert "A030-2" in items                                      # duplicate ref renamed
    assert items["A010"].duration_days is None and not items["A010"].predecessors   # summary cleaned
    assert items["A020"].duration_days == 0                        # milestone forced to 0 d
    assert items["A030"].predecessors[0].lag_days == 0             # lead removed
    reg = ex.scope_to_register(e.scope, doc)
    cit = dict(zip(reg["Ref"], reg["Citation"]))
    assert cit["A020"] == "Verified" and cit["A030"] == "Relocated" and cit["A040"] == "Unverified"
    assert reg.set_index("Ref").loc["A030", "Source"] == "P3"


def test_extract_output_limit_is_reported():
    doc = read_document(make_docx(), "sow.docx")
    with pytest.raises(ex.ExtractionError, match="too large"):
        ex.extract_scope(doc, "k", client=SimpleNamespace(models=_Models(_scope(), finish="MAX_TOKENS")))


def _error_client(fail: dict, scope, listed=None):
    """fail: {model: [codes to raise in order]}; afterwards the model succeeds. listed: model list or None."""
    from google.genai import errors
    calls = []

    class Models:
        def list(self):
            if listed is None:
                raise RuntimeError("listing not allowed")
            return [SimpleNamespace(name=f"models/{m}", supported_actions=["generateContent"]) for m in listed]

        def generate_content(self, model, contents, config):
            calls.append((model, config))
            queue = fail.get(model, [])
            if queue:
                code = queue.pop(0)
                exc = errors.ClientError.__new__(errors.ClientError)
                exc.code, exc.message, exc.status, exc.details = code, f"HTTP {code}", "", {}
                raise exc
            return SimpleNamespace(parsed=scope, text=None, usage_metadata=None,
                                   candidates=[SimpleNamespace(finish_reason=SimpleNamespace(name="STOP"))])
    return SimpleNamespace(models=Models()), calls


def test_model_fallback_on_404_like_a_new_free_tier_key(monkeypatch):
    """Field case 13-Oct-2026: gemini-2.5-flash returns 404 for a new free-tier key."""
    monkeypatch.setattr(ex.time, "sleep", lambda s: None)
    doc = read_document(make_docx(), "sow.docx")
    client, calls = _error_client({"gemini-2.5-flash": [404]}, _scope())
    e = ex.extract_scope(doc, "k", model="gemini-2.5-flash", client=client)
    assert [m for m, _ in calls] == ["gemini-2.5-flash", "gemini-3.8-flash"]
    assert e.model == "gemini-3.8-flash" and any("not available to this key" in n for n in e.notes)
    assert calls[0][1].temperature == 0.0 and calls[0][1].seed == 7        # Gemini 2.x rules on the first try


def test_listed_models_restrict_candidates_and_quota_falls_through(monkeypatch):
    monkeypatch.setattr(ex.time, "sleep", lambda s: None)
    doc = read_document(make_docx(), "sow.docx")
    client, calls = _error_client({"gemini-3.6-flash": [429, 429, 429]}, _scope(),
                                  listed=["gemini-3.6-flash", "gemini-3.5-flash-lite", "gemini-3.8-flash-tts"])
    e = ex.extract_scope(doc, "k", client=client)
    assert [m for m, _ in calls] == ["gemini-3.6-flash"] * 3 + ["gemini-3.5-flash-lite"]
    assert e.model == "gemini-3.5-flash-lite"


def test_unsupported_setting_400_steps_to_json_schema(monkeypatch):
    monkeypatch.setattr(ex.time, "sleep", lambda s: None)
    doc = read_document(make_docx(), "sow.docx")
    client, calls = _error_client({"gemini-3.8-flash": [400]}, _scope())
    e = ex.extract_scope(doc, "k", client=client)
    assert e.model == "gemini-3.8-flash"
    assert calls[0][1].response_schema is ex.ProjectScope and calls[1][1].response_json_schema is not None


def test_bad_key_and_all_models_missing_give_clear_messages(monkeypatch):
    monkeypatch.setattr(ex.time, "sleep", lambda s: None)
    doc = read_document(make_docx(), "sow.docx")
    client, _ = _error_client({"gemini-3.8-flash": [403]}, _scope())
    with pytest.raises(ex.ExtractionError, match="key was rejected"):
        ex.extract_scope(doc, "k", client=client)
    client, _ = _error_client({m: [404] for m in ex.MODEL_PREFERENCE}, _scope())
    with pytest.raises(ex.ExtractionError, match="None of these Gemini models"):
        ex.extract_scope(doc, "k", client=client)


def test_gemini_schema_has_no_defaults():
    pytest.importorskip("google.genai")
    from google import genai
    from google.genai import _transformers
    schema = _transformers.t_schema(genai.Client(api_key="offline-test")._api_client, ex.ProjectScope)
    dumped = schema.model_dump(exclude_none=True)
    assert dumped["properties"]["items"]["items"]["properties"]["item_type"]["enum"] == ["Summary", "Task",
                                                                                          "Milestone"]
    assert "default" not in str(dumped)


# --------------------------------------------------------------------------- #
# Review Gate validation
# --------------------------------------------------------------------------- #
def errors_of(rows, edit=None):
    df = register(rows)
    for (ref, col), val in (edit or {}).items():
        df.loc[df["Ref"] == ref, col] = val
    return [m for _, m in sb.validate_register(df).errors]


def test_clean_register_has_no_errors():
    assert errors_of(BASE) == []


@pytest.mark.parametrize("change, expect", [
    ({("A120", "Predecessors"): "A210"}, "Circular logic"),
    ({("A120", "Predecessors"): "A110 SF"}, "Start-to-Finish"),
    ({("A120", "Predecessors"): "A110 SS-2d"}, "Lead"),
    ({("A120", "Predecessors"): "A999"}, "not in the register"),
    ({("A120", "Predecessors"): "A120"}, "Lists itself"),
    ({("A100", "Predecessors"): "A030"}, "summary row"),
    ({("A120", "Duration (days)"): None}, "no duration"),
    ({("A120", "Ref"): "A110"}, "used more than once"),
    ({("A120", "WBS"): "7.1"}, "no parent row 7"),
    ({("A110", "Type"): "Summary"}, "no activities under it"),
    ({("A120", "Citation"): "Unverified"}, "Planner note"),
    ({("A030", "Include"): False}, "is excluded"),
])
def test_blocking_errors(change, expect):
    errs = errors_of(BASE, change)
    assert any(expect in e for e in errs), errs


def test_planner_note_clears_unverified_citation():
    assert errors_of(BASE, {("A120", "Citation"): "Unverified",
                              ("A120", "Planner note"): "Confirmed with client 10-Oct"}) == []


# --------------------------------------------------------------------------- #
# Register -> proven engine
# --------------------------------------------------------------------------- #
def test_build_zero_variance_and_policy():
    out = build(BASE)
    res, st = out.res, out.res.stats
    assert st["date_mismatches"] == 0 and st["finish_matches"]
    assert st["open_starts"] == 0 and st["open_ends"] == 0
    assert st["leads"] == 0
    assert res.nodes[0].kind == "project" and res.nodes[0].level == 1
    assert res.calendar_name == "Test Project CALENDAR"
    for n in res.nodes:
        if n.kind == "summary":
            assert n.children[-1].kind == "completion"
    # planner logic arrives unchanged: no link adjusted by the engine's fitting step
    assert not [i for i in res.issues if "Link from" in str(i[3]) and "changed" in str(i[3])]
    # single project finish: only the project completion milestone has no successor
    succ = {l.pred for n in res.nodes for l in n.preds}
    ends = [n for n in res.nodes if n.schedulable and n not in succ]
    assert [n.name for n in ends] == ["Test Project completed"]


def test_dates_are_deterministic_and_correct():
    res = build(BASE).res
    by = {n.name: cpm.display_dates(n, res.out_cal) for n in res.nodes}
    assert by["Contract award"] == (D0, D0)
    assert by["Mobilise team"] == (D0, dt.date(2026, 11, 13))
    assert by["Kick-off minutes issued"] == (D0, D0)                 # end of day 1, still dated 2 Nov
    assert by["Design basis"] == (dt.date(2026, 11, 16), dt.date(2026, 11, 27))
    assert by["Coating touch-up"][1] == by["Hydrotest"][1]          # FF honoured
    assert by["Handover"][0] == by["Coating touch-up"][1]            # FS from summary = after its last activity
    assert build(BASE).res.stats["output_finish"] == res.stats["output_finish"]


def test_summary_predecessor_resolves_to_terminal_activities():
    res = build(BASE).res
    hand = next(n for n in res.nodes if n.name == "Handover")
    assert {l.pred.name for l in hand.preds} >= {"Coating touch-up", "Kick-off minutes issued"}


def test_ff_that_would_start_before_its_predecessor_is_blocked():
    rows = [r if r[0] != "A230" else ("A230", "3.3", "Long coating", "Task", 20, "A220 FF") for r in BASE]
    with pytest.raises(sb.BuildError) as exc:
        build(rows)
    assert any("before A220 starts" in d for d in exc.value.details)


def test_holidays_and_six_day_week():
    res = build(BASE, work_week=6, holidays=[dt.date(2026, 11, 3)]).res
    assert res.stats["date_mismatches"] == 0 and res.info["work_week"]["days"] == 6
    mob = next(n for n in res.nodes if n.name == "Mobilise team")
    s, f = cpm.display_dates(mob, res.out_cal)
    assert (s, f) == (D0, dt.date(2026, 11, 13))      # 10 days Mon-Sat skipping Tue 3 Nov


def test_exports_open():
    res = build(BASE).res
    assert exporters.to_mspdi(res).startswith(b"<?xml")
    assert b"Test Project CALENDAR" in exporters.to_mspdi(res)
    assert exporters.to_csv(res).splitlines()[0].startswith(b"ID,Task Name")
    assert exporters.to_qa_xlsx(res)[:2] == b"PK"


# --------------------------------------------------------------------------- #
# DCMA-14 + audit trail
# --------------------------------------------------------------------------- #
def test_dcma14_on_clean_baseline():
    res = build(BASE).res
    checks = {c.no: c for c in dcma.dcma14(res)}
    assert len(checks) == 14
    for no in (1, 2, 3, 4, 5, 7, 12, 13):
        assert checks[no].status == "Pass", (no, checks[no])
    assert {checks[n].status for n in (9, 10, 11, 14)} == {"N/A"}
    assert dcma.dcma_xlsx(res, list(checks.values()))[:2] == b"PK"


def test_register_hash_is_stable_and_sensitive():
    a, b = register(BASE), register(list(reversed(BASE)))
    assert sb.register_hash(a) == sb.register_hash(b)            # row order does not matter
    b.loc[b["Ref"] == "A120", "Duration (days)"] = 31
    assert sb.register_hash(a) != sb.register_hash(b)            # any content change does
    assert sb.approval_record_xlsx(a, {"Reviewer": "T"})[:2] == b"PK"
