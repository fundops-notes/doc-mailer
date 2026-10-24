# -*- coding: utf-8 -*-
"""收件人「附件路径」链路 —— 导入 / 维护 / 流水线取路径 / 附件包 / 导出。

背景（用户实测反馈）：
  导入模板里有「附件路径」列，但界面上看不到、也没地方维护，
  流水线里更看不到「每个收件人各自的文件地址」是怎么配的。
  更严重的是：只要收件人配了附件路径，点「导出收件人清单」就直接崩
  （exe 是 --noconsole 构建的，异常被 tkinter 吞掉 → 点按钮毫无反应）。

本文件把这些点全部钉住。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

import lp_v3


HEADERS = ["收件人名称", "水印名称", "收件人", "抄送", "密送", "邮件标题",
           "附件路径", "打开口令", "权限口令"]


def _rec(name="甲公司", **kw):
    r = {"name": name, "to": "a@x.com", "open_pw": "open123", "owner_pw": "owner123"}
    r.update(kw)
    return lp_v3.normalize_recipient(r)


def _unit(rec):
    return {"recipients": [rec], "rep": rec, "lp_name": rec.get("name", ""),
            "group_key": rec.get("name", ""), "group": ""}


# ---------------------------------------------------------------- 导出

def test_export_recipients_with_attachment_paths_does_not_crash():
    """配了附件路径时导出不能崩（原来 attachments 被当 dict 用 → AttributeError）。"""
    recs = [_rec(attachments=[r"C:\资料\补充说明.pdf"])]
    rows = lp_v3.recipients_to_rows(recs)
    assert rows[1][6] == r"C:\资料\补充说明.pdf"


def test_export_recipients_accepts_dict_shaped_attachments():
    """原始数据里 attachments 可能是 dict 形态（老结构），也不能崩。"""
    rows = lp_v3.recipients_to_rows(
        [{"name": "甲公司", "to": "a@x.com", "attachments": [{"path": r"C:\资料\a.pdf"}]}])
    assert rows[1][6] == r"C:\资料\a.pdf"


def test_export_recipients_multiple_paths_joined():
    recs = [_rec(attachments=[r"C:\a.pdf", r"C:\b.pdf"])]
    rows = lp_v3.recipients_to_rows(recs)
    assert rows[1][6] == r"C:\a.pdf;C:\b.pdf"


# ---------------------------------------------------------------- 导入

def test_import_attachment_path_column_goes_to_custom_field():
    """v3.4.4：收件人页不再露出「附件路径」；导入该列只收进自定义字段，
    绝不写进 rec["attachments"] 主键。"""
    rows = [HEADERS, ["甲公司", "", "a@x.com", "", "", "",
                      r"C:\资料\补充说明.pdf;C:\资料\合同.pdf", "", ""]]
    recips, unknown, field_cols = lp_v3.parse_import_rows_v3(rows)
    norm = lp_v3.normalize_recipient(recips[0])
    assert norm["attachments"] == []
    assert "附件路径" in field_cols and "附件路径" in unknown
    assert norm["fields"]["附件路径"] == r"C:\资料\补充说明.pdf;C:\资料\合同.pdf"


def test_normalize_recipient_dedupes_files_and_attachments():
    """v2 的 files 与 v3 的 attachments 同时存在时不重复。"""
    rec = {"name": "甲公司", "attachments": [r"C:\a.pdf"],
           "files": r"C:\a.pdf;C:\b.pdf"}
    norm = lp_v3.normalize_recipient(rec)
    assert norm["attachments"] == [r"C:\a.pdf", r"C:\b.pdf"]


# ---------------------------------------------------------------- 路径来源

def test_normalize_item_keeps_path_source():
    it = lp_v3.normalize_item({"kind": "attach",
                               "params": {"path_source": "recipient.attachments"}})
    assert it["params"]["path_source"] == "recipient.attachments"


def test_normalize_item_defaults_path_source_to_manual():
    it = lp_v3.normalize_item({"kind": "attach"})
    assert it["params"]["path_source"] == lp_v3.PATH_SOURCE_MANUAL


def test_path_source_options_include_recipient_and_fields():
    rec = _rec(fields={"合同编号": "HT-001"})
    opts = lp_v3.path_source_options({}, rec, ["外部列A"])
    keys = [k for k, _g, _l in opts]
    assert lp_v3.PATH_SOURCE_MANUAL in keys
    assert lp_v3.PATH_SOURCE_RECIPIENT in keys
    assert "fields.合同编号" in keys
    assert "sheet.外部列A" in keys


def test_resolve_item_paths_manual():
    item = lp_v3.normalize_item({"kind": "attach", "params": {"files": [r"C:\a.pdf"]}})
    paths, label = lp_v3.resolve_item_paths(item, _unit(_rec()), {}, None)
    assert paths == [r"C:\a.pdf"]
    assert label == "手填路径"


def test_resolve_item_paths_from_recipient_attachments():
    rec = _rec(attachments=[r"C:\每人\甲.pdf"])
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "recipient.attachments"}})
    paths, label = lp_v3.resolve_item_paths(item, _unit(rec), {}, None)
    assert paths == [r"C:\每人\甲.pdf"]
    assert "附件路径" in label


def test_resolve_item_paths_from_recipient_field():
    rec = _rec(fields={"合同路径": r"C:\每人\合同.pdf"})
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "fields.合同路径"}})
    paths, _label = lp_v3.resolve_item_paths(item, _unit(rec), {}, None)
    assert paths == [r"C:\每人\合同.pdf"]


def test_resolve_item_paths_splits_multiple_values():
    rec = _rec(fields={"合同路径": r"C:\a.pdf;C:\b.pdf"})
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "fields.合同路径"}})
    paths, _label = lp_v3.resolve_item_paths(item, _unit(rec), {}, None)
    assert paths == [r"C:\a.pdf", r"C:\b.pdf"]


# ---------------------------------------------------------------- 执行

def test_execute_item_attach_uses_recipient_paths(tmp_path):
    src = tmp_path / "补充说明.pdf"
    src.write_bytes(b"%PDF-1.4 test")
    rec = _rec(attachments=[str(src)])
    item = lp_v3.normalize_item({"kind": "attach", "name": "补充材料",
                                  "params": {"path_source": "recipient.attachments"}})
    event = {"options": {}, "mail": {}, "recipients": [rec], "pipeline": [item]}
    outdir = tmp_path / "out"
    outdir.mkdir()
    ctx = lp_v3.RunContext(event, outdir)
    files, warns = lp_v3.execute_item(item, _unit(rec), event, ctx)
    assert len(files) == 1
    assert Path(files[0]).exists()
    assert Path(files[0]).read_bytes() == b"%PDF-1.4 test"


def test_execute_item_attach_missing_recipient_path_is_warning_not_error(tmp_path):
    """按收件人取路径时某人没填，是正常情况：跳过并提醒，不该让整条失败。"""
    rec = _rec(attachments=[str(tmp_path / "没有这个文件.pdf")])
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "recipient.attachments"}})
    event = {"options": {}, "mail": {}, "recipients": [rec], "pipeline": [item]}
    outdir = tmp_path / "out"
    outdir.mkdir()
    ctx = lp_v3.RunContext(event, outdir)
    files, warns = lp_v3.execute_item(item, _unit(rec), event, ctx)
    assert files == []
    assert any("不存在" in w for w in warns)


def test_execute_item_attach_no_recipient_path_is_warning_not_error(tmp_path):
    rec = _rec()          # 没填附件路径
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "recipient.attachments"}})
    event = {"options": {}, "mail": {}, "recipients": [rec], "pipeline": [item]}
    outdir = tmp_path / "out"
    outdir.mkdir()
    ctx = lp_v3.RunContext(event, outdir)
    files, warns = lp_v3.execute_item(item, _unit(rec), event, ctx)
    assert files == []
    assert warns


def test_execute_item_attach_manual_missing_file_still_fails(tmp_path):
    """手填路径不存在仍然是配置错误（保持旧行为：整条失败）。"""
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"files": [str(tmp_path / "没有.pdf")]}})
    event = {"options": {}, "mail": {}, "recipients": [], "pipeline": [item]}
    outdir = tmp_path / "out"
    outdir.mkdir()
    ctx = lp_v3.RunContext(event, outdir)
    with pytest.raises(RuntimeError):
        lp_v3.execute_item(item, _unit(_rec()), event, ctx)


# ---------------------------------------------------------------- 附件包组装

def test_event_uses_recipient_paths_detects_item():
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "recipient.attachments"}})
    assert lp_v3.event_uses_recipient_paths({"pipeline": [item]}) is True
    assert lp_v3.event_uses_recipient_paths({"pipeline": []}) is False


def test_assemble_attachments_no_duplicate_when_item_uses_recipient_paths(tmp_path):
    """条目已经按路径来源产出时，收件人那列路径不能再原样附加一次。"""
    src = tmp_path / "补充说明.pdf"
    src.write_bytes(b"%PDF-1.4 test")
    rec = _rec(attachments=[str(src)])
    item = lp_v3.normalize_item({"kind": "attach",
                                  "params": {"path_source": "recipient.attachments"}})
    event = {"options": {}, "mail": {}, "recipients": [rec], "pipeline": [item],
             "attachments": []}
    produced = tmp_path / "out" / "补充材料_甲公司.pdf"
    produced.parent.mkdir(parents=True, exist_ok=True)
    produced.write_bytes(b"%PDF-1.4 test")
    files, _missing = lp_v3.assemble_attachments(event, _unit(rec), [produced])
    paths = [str(f) for f in files]
    assert str(src) not in paths       # 原始路径不再重复附加
    assert str(produced) in paths      # 条目产物在


def test_assemble_attachments_still_adds_recipient_paths_by_default(tmp_path):
    """没有条目用路径来源时，收件人的附件路径照旧直接进附件包。"""
    src = tmp_path / "补充说明.pdf"
    src.write_bytes(b"%PDF-1.4 test")
    rec = _rec(attachments=[str(src)])
    event = {"options": {}, "mail": {}, "recipients": [rec], "pipeline": [],
             "attachments": []}
    files, _missing = lp_v3.assemble_attachments(event, _unit(rec), [])
    assert str(src) in [str(f) for f in files]


# ---------------------------------------------------------------- 预检

def test_preflight_warns_missing_recipient_attachment(tmp_path):
    rec = _rec(attachments=[str(tmp_path / "没有.pdf")])
    event = {"recipients": [rec], "pipeline": [], "attachments": [],
             "options": {}, "mail": {}}
    errors, warnings, _info = lp_v3.preflight(event, None)
    assert any("专属附件不存在" in w for w in warnings)
