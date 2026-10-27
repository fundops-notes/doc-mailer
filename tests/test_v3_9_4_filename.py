"""Attachment naming, protected fields and per-recipient output isolation."""
from pathlib import Path
import sys
import tkinter as tk
from tkinter import ttk

import pytest
import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import lp_v3
import lp_tool


def unit(rec):
    return {"rep": rec, "recipients": [rec], "lp_name": rec["name"]}


def item(files, **params):
    return lp_v3.normalize_item({"id": "file", "kind": "attach",
                                 "params": {"files": list(map(str, files)), **params}})


def test_attach_template_fields_and_legacy_preserve(tmp_path):
    source = tmp_path / "report.pdf"
    source.write_bytes(b"report")
    rec = {"name": "Alice", "wm": "Mark", "open_pw": "open123", "owner_pw": "secret123"}
    event = {"recipients": [rec], "options": {}}
    ctx = lp_v3.RunContext(event, tmp_path / "out")
    old, _ = lp_v3.execute_item(item([source], preserve_name=True), unit(rec), event, ctx)
    new, _ = lp_v3.execute_item(item([source], preserve_name=True, filename_tpl="{{原件名}}_{{水印名称}}_{{打开口令}}_{{收件人名称}}"), unit(rec), event, ctx)
    assert old[0].name == "report.pdf"
    assert new[0].name == "report_Mark_open123_Alice.pdf"
    assert new[0].read_bytes() == b"report"


@pytest.mark.parametrize("template", ["{{权限口令}}", "{{owner_pw}}", "report_secret123"])
def test_owner_secret_never_becomes_attachment_name(tmp_path, template):
    source = tmp_path / "report.pdf"
    source.write_bytes(b"report")
    rec = {"name": "Alice", "owner_pw": "secret123", "open_pw": "open123"}
    event = {"recipients": [rec], "options": {}}
    with pytest.raises(ValueError, match="权限口令"):
        lp_v3.execute_item(item([source], filename_tpl=template), unit(rec), event, lp_v3.RunContext(event, tmp_path / "out"))


def test_custom_attachment_names_cannot_contain_any_event_owner_password():
    event = {"recipients": [{"owner_pw": "other-owner"}], "options": {"gp_owner_pw": "gp-secret"}}
    for name in ("report_other-owner.pdf", "report_gp-secret.pdf"):
        with pytest.raises(ValueError, match="权限口令"):
            lp_v3._check_attachment_name(name, event)


def test_duplicate_names_cannot_overwrite_other_recipient(tmp_path):
    a = tmp_path / "a" / "same.pdf"; a.parent.mkdir(); a.write_bytes(b"alice")
    b = tmp_path / "b" / "same.pdf"; b.parent.mkdir(); b.write_bytes(b"bob")
    ra = {"name": "Alice", "attachments": [str(a)]}
    rb = {"name": "Bob", "attachments": [str(b)]}
    event = {"recipients": [ra, rb], "options": {}}
    it = item([], path_source="recipient.attachments", filename_tpl="same")
    ctx = lp_v3.RunContext(event, tmp_path / "out")
    oa, _ = lp_v3.execute_item(it, unit(ra), event, ctx)
    ob, _ = lp_v3.execute_item(it, unit(rb), event, ctx)
    assert oa[0] != ob[0]
    assert (oa[0].read_bytes(), ob[0].read_bytes()) == (b"alice", b"bob")


def test_partial_rerun_preserves_previous_recipient_output(tmp_path):
    source = tmp_path / "source.pdf"; source.write_bytes(b"new")
    output = tmp_path / "out"; old = output / "file" / "same.pdf"
    old.parent.mkdir(parents=True); old.write_bytes(b"other recipient")
    a = {"name": "Alice"}; b = {"name": "Bob"}
    it = item([source], filename_tpl="same")
    event = {"recipients": [a, b], "pipeline": [it], "process_results": {"Bob": [str(old)]}, "options": {}}
    outputs, _, errors = lp_v3.run_items(event, output, only_recips=["Alice"])
    assert not errors
    new = next(iter(outputs[it["id"]].values()))[0]
    assert new != old and new.read_bytes() == b"new" and old.read_bytes() == b"other recipient"


def test_pdf_conversion_reserves_final_name(tmp_path, monkeypatch):
    doc = tmp_path / "report.txt"; doc.write_text("document")
    pdf = tmp_path / "report.pdf"; pdf.write_bytes(b"pdf")
    rec = {"name": "Alice"}; event = {"recipients": [rec], "options": {}}
    monkeypatch.setattr(lp_v3.wm, "file_to_pdf", lambda src, dst, batch=None: Path(dst).write_bytes(Path(src).read_bytes()))
    files, _ = lp_v3.execute_item(item([doc, pdf], filename_tpl="same", to_pdf=True), unit(rec), event, lp_v3.RunContext(event, tmp_path / "out"))
    assert len({p.name for p in files}) == 2
    assert [p.read_bytes() for p in files] == [b"document", b"pdf"]


def test_real_pdf_watermark_and_encryption_named_per_recipient(tmp_path):
    source = tmp_path / "report.pdf"
    doc = fitz.open(); page = doc.new_page(); page.insert_text((72, 72), "Quarterly report")
    doc.save(str(source)); doc.close()
    rec = {"name": "Alice", "wm": "Alice mark", "open_pw": "open123", "owner_pw": "private456"}
    event = {"recipients": [rec], "options": {}}
    it = item([source], filename_tpl="{{原件名}}_{{收件人名称}}_{{打开口令}}", watermark=True, encrypt=True)
    files, _ = lp_v3.execute_item(it, unit(rec), event, lp_v3.RunContext(event, tmp_path / "out"))
    assert files[0].name == "report_Alice_open123.pdf"
    result = fitz.open(str(files[0]))
    assert result.needs_pass and result.authenticate("open123")
    assert "Quarterly report" in result[0].get_text()
    result.close()


def test_outlook_draft_attachment_is_checked_before_save(tmp_path, monkeypatch):
    import win32com.client
    class Attachments:
        def __init__(self): self.paths = []
        def Add(self, path): self.paths.append(path)
    class Mail:
        def __init__(self): self.Attachments = Attachments(); self.saved = False
        def Save(self): self.saved = True
    mail = Mail()
    class Outlook:
        def CreateItem(self, kind): return mail
    monkeypatch.setattr(win32com.client, "Dispatch", lambda _: Outlook())
    safe = tmp_path / "report_open123.pdf"; safe.write_bytes(b"pdf")
    lp_v3.create_outlook_draft({"to": "a@example.com"}, [safe], "subject", "body", owner_pw="private456")
    assert mail.saved and Path(mail.Attachments.paths[0]).name == safe.name
    with pytest.raises(ValueError, match="权限口令"):
        lp_v3.create_outlook_draft({}, [tmp_path / "report_private456.pdf"], "s", "b", owner_pw="private456")


@pytest.mark.parametrize("variant", ["pdf", "file", "excel"])
def test_attachment_editor_inserts_and_flushes_template(tmp_path, monkeypatch, tk_root, variant):
    monkeypatch.setattr(lp_v3, "default_base_dir", lambda: tmp_path)
    app = lp_tool.App(tk_root)
    try:
        it = lp_v3.normalize_item({"id": "test", "kind": "attach", "params": {"input_type": variant}})
        app._wb_event = {"pipeline": [it], "recipients": [], "options": {}}
        parent = ttk.Frame(tk_root); parent.pack()
        app._wb_item_params(parent, it)
        def walk(w):
            yield w
            for child in w.winfo_children():
                yield from walk(child)
        button = next(w for w in walk(parent) if isinstance(w, ttk.Menubutton) and w.cget("text") == "插入字段")
        menu = button.nametowidget(button.cget("menu"))
        menu.invoke(0); menu.invoke(3)
        assert app._wb_flush_item_editor()
        assert it["params"]["filename_tpl"] == "{{原件名}}{{收件人名称}}"
        parent.destroy()
    finally:
        app.shutdown()
