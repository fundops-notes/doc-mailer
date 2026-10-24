# -*- coding: utf-8 -*-
"""v3 执行引擎回归测试（预检 / 附件组装 / 草稿 / 执行记录 / 只重跑失败）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz
import pytest

from lp_v3 import (Store, assemble_attachments, build_field_map, create_outlook_draft,
                   draft_units, failed_recipients, new_event, normalize_event,
                   normalize_item, normalize_recipient, preflight, preview_plan,
                   render_with_fields, run_event)


def _event(recips=None, pipeline=None, **opts):
    ev = new_event("E")
    ev["recipients"] = [normalize_recipient(r) for r in (recips or [])]
    if pipeline is not None:
        ev["pipeline"] = [normalize_item(s) for s in pipeline]
    ev["options"].update(opts)
    return ev


def _pdf(path, text="hello"):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text)
    doc.save(str(path))
    doc.close()
    return path


# ---------------------------------------------------------------- 预检

def test_preflight_blocks_on_empty_recipients():
    errors, warnings, info = preflight(_event())
    assert any("收件人清单为空" in e for e in errors)
    assert info["recipients"] == 0


def test_preflight_blocks_missing_and_duplicate_passwords():
    ev = _event(recips=[
        {"name": "甲", "to": "a@x.com", "open_pw": "", "owner_pw": "ow1"},
        {"name": "乙", "to": "b@x.com", "open_pw": "same", "owner_pw": "same"},
        {"name": "丙", "to": "c@x.com", "open_pw": "dup", "owner_pw": "ow3"},
        {"name": "丁", "to": "d@x.com", "open_pw": "dup", "owner_pw": "ow4"},
    ])
    errors, _w, _i = preflight(ev)
    assert any("打开口令为空" in e for e in errors)
    assert any("相同" in e for e in errors)
    assert any("重复" in e for e in errors)


def test_preflight_gp_password_conflict():
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "gppw",
                         "owner_pw": "ow1"}],
                gp_owner_mode=True, gp_owner_pw="gppw")
    errors, _w, _i = preflight(ev)
    assert any("GP 统一权限口令" in e for e in errors)


def test_preflight_missing_attachment_is_warning_not_error(tmp_path):
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "op1",
                         "owner_pw": "ow1"}])
    ev["attachments"] = [{"name": "缺失", "path": str(tmp_path / "nope.pdf"),
                          "scope": "all", "targets": [], "group": ""}]
    errors, warnings, _i = preflight(ev)
    assert not any("附件不存在" in e for e in errors)
    assert any("附件不存在" in w for w in warnings)


def test_preflight_blocks_missing_template_file(tmp_path):
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "op1",
                         "owner_pw": "ow1"}],
                pipeline=[{"kind": "generate",
                           "params": {"template_path": str(tmp_path / "nope.docx")}}])
    errors, _w, _i = preflight(ev)
    assert any("Word 模板不存在" in e for e in errors)


def test_preflight_blocks_empty_pipeline():
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "op1",
                         "owner_pw": "ow1"}])
    errors, _w, _i = preflight(ev)
    assert any("流水线为空" in e for e in errors)


def test_preflight_detects_dangling_placeholder_in_real_docx(word_template):
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "op1",
                         "owner_pw": "ow1"}],
                pipeline=[{"kind": "generate",
                           "params": {"template_path": str(word_template)}}])
    # 模板里有 {{LP_Name}} {{Call_Amount}} {{Due_Date}}，都没给收件人字段
    errors, _w, _i = preflight(ev)
    assert any("悬空占位符" in e for e in errors)


def test_preflight_pass_with_mapped_placeholders(word_template):
    ev = _event(recips=[{"name": "甲公司", "to": "a@x.com", "open_pw": "op1",
                         "owner_pw": "ow1",
                         "fields": {"Call_Amount": "200,000",
                                    "Due_Date": "2026-11-15"}}],
                pipeline=[{"kind": "generate",
                           "params": {"template_path": str(word_template)}}])
    errors, _w, _i = preflight(ev)
    assert not [e for e in errors if "悬空" in e]


# ---------------------------------------------------------------- 附件组装

def test_assemble_attachments_respects_scope(tmp_path):
    a_all = tmp_path / "all.pdf"
    a_all.write_bytes(b"content-all")
    a_one = tmp_path / "one.pdf"
    a_one.write_bytes(b"content-one")   # 内容各不相同（同内容会被哈希去重合并）
    a_grp = tmp_path / "grp.pdf"
    a_grp.write_bytes(b"content-grp")
    ev = _event(recips=[{"name": "甲", "group": "G1"}, {"name": "乙", "group": "G2"}])
    ev["attachments"] = [
        {"name": "全员", "path": str(a_all), "scope": "all", "targets": [], "group": ""},
        {"name": "只给甲", "path": str(a_one), "scope": "specific",
         "targets": ["甲"], "group": ""},
        {"name": "G1组", "path": str(a_grp), "scope": "group", "targets": [], "group": "G1"},
    ]
    unit = dict(draft_units(ev))["G1"]
    files, missing = assemble_attachments(ev, unit, [])
    names = sorted(p.name for p in files)
    assert names == ["all.pdf", "grp.pdf", "one.pdf"]
    assert missing == []


def test_assemble_attachments_reports_missing(tmp_path):
    ev = _event(recips=[{"name": "甲"}])
    ev["attachments"] = [{"name": "缺", "path": str(tmp_path / "nope.pdf"),
                          "scope": "all", "targets": [], "group": ""}]
    unit = dict(draft_units(ev))["§甲"]
    files, missing = assemble_attachments(ev, unit, [])
    assert files == []
    assert len(missing) == 1


def test_draft_units_merges_group():
    ev = _event(recips=[{"name": "甲", "group": "G1"}, {"name": "乙", "group": "G1"},
                        {"name": "丙"}])
    keys = [k for k, _u in draft_units(ev)]
    assert "G1" in keys
    assert len(keys) == 2          # G1 一组 + 丙自己一组
    ev["options"]["send_mode"] = "per_lp"
    assert len(draft_units(ev)) == 3


def test_draft_units_legacy_merge_flag_still_works():
    """v2 迁移来的老数据只有 merge_groups 布尔，没有 send_mode。"""
    ev = {"options": {"merge_groups": False},
          "recipients": [{"name": "甲", "group": "G1"},
                         {"name": "乙", "group": "G1"}, {"name": "丙"}]}
    assert len(draft_units(ev)) == 3
    ev["options"]["merge_groups"] = True
    assert len(draft_units(ev)) == 2


# ---------------------------------------------------------------- 草稿（Mock Outlook）

class _FakeAttachments:
    def __init__(self, mail):
        self.mail = mail

    def Add(self, path):    # noqa: N802
        self.mail.attachments.append(path)


class _FakeMail:
    def __init__(self):
        self.To = self.CC = self.BCC = self.Subject = self.HTMLBody = ""
        self.SentOnBehalfOfName = ""
        self.attachments = []
        self.saved = 0
        self.sent = 0

    @property
    def Attachments(self):  # noqa: N802 - COM 风格：属性返回集合
        return _FakeAttachments(self)

    def Save(self):         # noqa: N802
        self.saved += 1

    def Send(self):         # noqa: N802
        self.sent += 1


class _FakeOutlook:
    def __init__(self):
        self.mails = []

    def CreateItem(self, _kind):  # noqa: N802
        m = _FakeMail()
        self.mails.append(m)
        return m


@pytest.fixture()
def fake_outlook(monkeypatch):
    fake = _FakeOutlook()
    import win32com.client as win32
    monkeypatch.setattr(win32, "Dispatch", lambda *_a, **_k: fake)
    return fake


def test_draft_never_sends_and_scrubs_owner_password(fake_outlook):
    rec = {"name": "甲公司", "to": "a@x.com", "cc": "c@x.com", "bcc": "",
           "open_pw": "Open123456", "owner_pw": "OwnerSecret123"}
    fmap = build_field_map(rec, _event())
    body = "<p>口令 {{打开口令}} 密 {{密}}</p>"
    create_outlook_draft(rec, [], "主题 {{收件人名称}}", body, fmap,
                         owner_pw="OwnerSecret123")
    m = fake_outlook.mails[0]
    assert m.saved == 1
    assert m.sent == 0, "绝不允许出现 Send 调用"
    assert "OwnerSecret123" not in m.HTMLBody
    assert "OwnerSecret123" not in m.Subject
    assert "Open123456" in m.HTMLBody
    assert m.To == "a@x.com" and m.CC == "c@x.com"


def test_draft_merges_group_addresses(fake_outlook):
    members = [{"name": "甲", "to": "a@x.com", "cc": "cc1@x.com", "bcc": "",
                "open_pw": "op1", "owner_pw": "ow1"},
               {"name": "乙", "to": "b@x.com", "cc": "cc1@x.com", "bcc": "bcc@x.com",
                "open_pw": "op2", "owner_pw": "ow2"}]
    rec = dict(members[0])
    rec["to"] = "a@x.com;b@x.com"
    rec["cc"] = "cc1@x.com"
    rec["bcc"] = "bcc@x.com"
    create_outlook_draft(rec, [], "s", "b", build_field_map(members[0], _event()))
    m = fake_outlook.mails[0]
    assert m.To == "a@x.com;b@x.com"
    assert m.CC == "cc1@x.com"     # 去重
    assert m.BCC == "bcc@x.com"


# ---------------------------------------------------------------- 执行记录

def test_run_event_end_to_end(tmp_path):
    src = _pdf(tmp_path / "src.pdf")
    store = Store(tmp_path)
    store.ensure_dirs()
    ev = _event(recips=[{"name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
                         "owner_pw": "Owner123456", "wm": "甲公司水印"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)],
                                      "watermark": True, "encrypt": True}}])
    ev["options"]["outdir"] = str(tmp_path / "out")
    run = run_event(store, ev, create_drafts=False, outdir=str(tmp_path / "out"))
    assert run["status"] == "ok", run
    assert run["stats"]["ok"] == 1
    item = run["items"][0]
    assert item["files"], "应该产出文件"
    assert Path(item["files"][0]).exists()
    assert item["open_pw_fp"] and len(item["open_pw_fp"]) == 8
    # 产出必须是加密的
    d = fitz.open(item["files"][0])
    try:
        assert d.needs_pass
    finally:
        d.close()
    # 执行记录落盘
    runs = store.list_runs(ev["id"])
    assert len(runs) == 1


def test_run_event_preflight_failure_produces_nothing(tmp_path):
    store = Store(tmp_path)
    store.ensure_dirs()
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "", "owner_pw": ""}])
    run = run_event(store, ev, create_drafts=False, outdir=str(tmp_path / "out"))
    assert run["status"] == "preflight_failed"
    assert run["items"] == []
    assert run["errors"]


def test_run_event_creates_drafts(fake_outlook, tmp_path):
    src = _pdf(tmp_path / "src.pdf")
    store = Store(tmp_path)
    store.ensure_dirs()
    ev = _event(recips=[{"name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
                         "owner_pw": "Owner123456"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)], "encrypt": True}}])
    ev["options"]["outdir"] = str(tmp_path / "out")
    ev["mail"]["subject"] = "机密文件（{{收件人名称}} 专属）"
    ev["mail"]["body"] = "<p>打开口令：{{打开口令}}</p>"
    run = run_event(store, ev, create_drafts=True, outdir=str(tmp_path / "out"))
    assert len(fake_outlook.mails) == 1
    m = fake_outlook.mails[0]
    assert m.sent == 0
    assert "Owner123456" not in m.HTMLBody
    assert "Open123456" in m.HTMLBody
    assert m.attachments, "草稿应带上附件"
    assert run["stats"]["ok"] == 1


def test_merged_group_attaches_every_members_file(tmp_path):
    """合并组：组内每人的产物都要进同一个附件包。"""
    src = _pdf(tmp_path / "src.pdf")
    store = Store(tmp_path)
    store.ensure_dirs()
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "group": "G1",
                         "open_pw": "op12345678", "owner_pw": "ow12345678"},
                        {"name": "乙", "to": "b@x.com", "group": "G1",
                         "open_pw": "op87654321", "owner_pw": "ow87654321"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)], "encrypt": True}}])
    ev["options"]["outdir"] = str(tmp_path / "out")
    run = run_event(store, ev, create_drafts=False, outdir=str(tmp_path / "out"))
    assert len(run["items"]) == 1, "同组应该只有一封"
    item = run["items"][0]
    assert item["members"] == ["甲", "乙"]
    assert len(item["files"]) == 2, "组内两人的文件都要进附件包：%s" % item["files"]
    names = sorted(Path(f).name for f in item["files"])
    assert any("甲" in n for n in names) and any("乙" in n for n in names)


def test_all_item_outputs_go_into_attachment_package(tmp_path):
    """回归：附件包里必须有加密件。

    v3.1 规则：所有条目的产物都进附件包（每个条目内部已把中间产物处理掉）。
    """
    src = _pdf(tmp_path / "src.pdf")
    note = tmp_path / "note.txt"
    note.write_text("hello", encoding="utf-8")
    store = Store(tmp_path)
    store.ensure_dirs()
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "op12345678",
                         "owner_pw": "ow12345678"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)], "encrypt": True}},
                          {"kind": "attach",
                           "params": {"files": [str(note)]}}])
    ev["options"]["outdir"] = str(tmp_path / "out")
    run = run_event(store, ev, create_drafts=False, outdir=str(tmp_path / "out"))
    assert run["status"] == "ok", run
    files = run["items"][0]["files"]
    assert len(files) == 2, "加密 PDF + 静态文件都该在附件里：%s" % files
    pdfs = [f for f in files if f.endswith(".pdf")]
    assert pdfs, "加密件丢了：%s" % files
    d = fitz.open(pdfs[0])
    try:
        assert d.needs_pass, "附件里的 PDF 必须仍是加密的"
    finally:
        d.close()
    assert any(f.endswith(".txt") for f in files), files


def test_preview_plan_lists_all_members_for_merged_group(tmp_path):
    src = _pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "group": "G1"},
                        {"name": "乙", "to": "b@x.com", "group": "G1"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)], "encrypt": True}}])
    plan = preview_plan(ev)
    assert len(plan) == 1
    _key, _unit, names = plan[0]
    assert len(names) == 2, names


def test_failed_recipients_extraction():
    run = {"items": [
        {"members": ["甲"], "draft": "ok", "error": ""},
        {"members": ["乙", "丙"], "draft": "失败:xxx", "error": "boom"},
        {"members": ["丁"], "draft": "ok", "error": "文件缺失"},
    ]}
    assert failed_recipients(run) == {"乙", "丙", "丁"}


def test_rerun_only_failed_recipients(tmp_path):
    src = _pdf(tmp_path / "src.pdf")
    store = Store(tmp_path)
    store.ensure_dirs()
    ev = _event(recips=[{"name": "甲", "to": "a@x.com", "open_pw": "op12345678",
                         "owner_pw": "ow12345678"},
                        {"name": "乙", "to": "b@x.com", "open_pw": "op87654321",
                         "owner_pw": "ow87654321"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)], "encrypt": True}}])
    ev["options"]["outdir"] = str(tmp_path / "out")
    run = run_event(store, ev, create_drafts=False, outdir=str(tmp_path / "out"))
    assert len(run["items"]) == 2
    run2 = run_event(store, ev, only_recips={"乙"}, create_drafts=False,
                     outdir=str(tmp_path / "out"))
    assert len(run2["items"]) == 1
    assert run2["items"][0]["members"] == ["乙"]


def test_preview_plan_lists_expected_files(tmp_path):
    src = _pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲公司", "to": "a@x.com"}],
                pipeline=[{"kind": "attach",
                           "params": {"files": [str(src)], "encrypt": True}}])
    plan = preview_plan(ev)
    assert len(plan) == 1
    key, unit, names = plan[0]
    assert names and names[0].endswith(".pdf")
    assert "甲公司" in names[0]


# ---------------------------------------------------------------- 静态铁律

def test_no_send_call_anywhere_in_source():
    """铁律：全代码库不存在任何 Send 调用。"""
    src_dir = Path(__file__).resolve().parents[1] / "src"
    for name in ("lp_tool.py", "lp_v3.py", "watermark_encrypt.py"):
        text = (src_dir / name).read_text(encoding="utf-8")
        assert ".Send(" not in text, "%s 里出现了 Send 调用" % name
        assert ".Send " not in text, "%s 里出现了 Send 调用" % name


def test_owner_password_never_in_template_field_table():
    """铁律：模板字段表里没有权限口令。"""
    ev = _event()
    rec = normalize_recipient({"name": "甲", "open_pw": "op", "owner_pw": "SECRET"})
    fmap = build_field_map(rec, ev)
    for k, v in fmap.items():
        assert "SECRET" not in str(v)
        assert "owner" not in k.lower()
        assert "权限口令" not in k
