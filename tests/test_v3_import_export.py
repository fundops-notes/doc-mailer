# -*- coding: utf-8 -*-
"""导入校验 / 导出备份 / 发送模式 / 流程图 —— 纯逻辑测试（不需要 GUI）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

import lp_v3


HEADERS = ["收件人名称", "收件人", "抄送", "打开口令", "权限口令", "Call_Amount"]


def _rows(*data):
    return [HEADERS] + [list(r) for r in data]


def _rec(name, to="a@x.com", **kw):
    r = {"name": name, "to": to, "open_pw": "open123", "owner_pw": "owner123"}
    r.update(kw)
    return lp_v3.normalize_recipient(r)


# ---------------------------------------------------------------- 导入校验

def test_validate_empty_file_is_fatal():
    res = lp_v3.validate_recipient_rows([])
    assert res["ok"] is False
    assert res["fatal"]


def test_validate_missing_name_column_is_fatal():
    res = lp_v3.validate_recipient_rows([["随便一列", "值"]])
    assert res["ok"] is False
    assert "收件人名称" in res["fatal"][0]


def test_validate_ok_row():
    res = lp_v3.validate_recipient_rows(
        _rows(["甲公司", "a@x.com", "", "open123", "owner123", "200,000"]))
    assert res["ok"] is True
    assert len(res["items"]) == 1
    assert res["items"][0]["status"] == "ok"
    assert res["items"][0]["no"] == 2          # 2 = 表格第 2 行（第 1 行是表头）


def test_validate_empty_email_is_error():
    res = lp_v3.validate_recipient_rows(_rows(["甲公司", "", "", "", "", ""]))
    assert res["items"][0]["status"] == "error"
    assert any("邮箱为空" in e for e in res["items"][0]["errors"])


def test_validate_bad_email_format_is_error():
    res = lp_v3.validate_recipient_rows(_rows(["甲公司", "not-an-email", "", "", "", ""]))
    assert res["items"][0]["status"] == "error"
    assert any("格式不对" in e for e in res["items"][0]["errors"])


def test_validate_cc_email_format_is_error():
    res = lp_v3.validate_recipient_rows(
        _rows(["甲公司", "a@x.com", "bad@@x.com", "", "", ""]))
    assert res["items"][0]["status"] == "error"
    assert any("抄送" in e for e in res["items"][0]["errors"])


def test_validate_duplicate_name_in_file_is_error():
    res = lp_v3.validate_recipient_rows(
        _rows(["甲公司", "a@x.com", "", "", "", ""],
              ["甲公司", "b@x.com", "", "", "", ""]))
    assert res["items"][1]["status"] == "error"
    assert any("重名" in e for e in res["items"][1]["errors"])
    assert res["duplicate_names"] == ["甲公司"]


def test_validate_existing_name_is_warning():
    res = lp_v3.validate_recipient_rows(
        _rows(["甲公司", "a@x.com", "", "", "", ""]), existing_names=["甲公司"])
    assert res["items"][0]["status"] == "warning"
    assert any("同名" in w for w in res["items"][0]["warnings"])


def test_validate_same_open_and_owner_password_is_error():
    """铁律：打开口令 ≠ 权限口令。"""
    res = lp_v3.validate_recipient_rows(
        _rows(["甲公司", "a@x.com", "", "samepass", "samepass", ""]))
    assert res["items"][0]["status"] == "error"
    assert any("打开口令与权限口令相同" in e for e in res["items"][0]["errors"])


def test_validate_unknown_column_becomes_custom_field():
    res = lp_v3.validate_recipient_rows(
        _rows(["甲公司", "a@x.com", "", "", "", "200,000"]))
    assert res["field_cols"] == ["Call_Amount"]
    assert res["unknown_cols"] == ["Call_Amount"]     # 未识别 → 当作自定义字段收进来
    assert res["items"][0]["record"]["fields"]["Call_Amount"] == "200,000"


def test_validate_missing_attachment_file_is_warning():
    head = ["收件人名称", "收件人", "附件路径"]
    res = lp_v3.validate_recipient_rows(
        [head, ["甲公司", "a@x.com", "C:/不存在/文件.pdf"]])
    assert res["items"][0]["status"] == "warning"
    assert any("不存在" in w for w in res["items"][0]["warnings"])


def test_validate_example_placeholder_row_is_skipped():
    res = lp_v3.validate_recipient_rows(
        _rows(["示例：苏州某某科技有限公司", "a@x.com", "", "", "", ""]))
    assert res["items"] == []


# ---------------------------------------------------------------- 导出

def test_recipients_to_rows_roundtrip():
    recips = [
        _rec("甲公司", "a@x.com", cc="c@x.com", group="G1",
             fields={"Call_Amount": "200,000"}),
        _rec("乙公司", "b@x.com", fields={"Unfunded": "800,000"}),
    ]
    rows = lp_v3.recipients_to_rows(recips)
    assert rows[0][:5] == ["收件人名称", "水印名称", "收件人", "抄送", "密送"]
    assert "Call_Amount" in rows[0] and "Unfunded" in rows[0]
    back, _unknown, _fc = lp_v3.parse_import_rows_v3(rows)
    assert [r["name"] for r in back] == ["甲公司", "乙公司"]
    assert back[0]["group"] == "G1"
    assert back[0]["fields"]["Call_Amount"] == "200,000"
    assert back[1]["fields"]["Unfunded"] == "800,000"


def test_recipients_to_rows_can_hide_passwords():
    rows = lp_v3.recipients_to_rows([_rec("甲公司")], include_passwords=False)
    head = rows[0]
    assert rows[1][head.index("打开口令")] == ""
    assert rows[1][head.index("权限口令")] == ""


def test_event_export_import_roundtrip():
    ev = lp_v3.new_event("备份事件")
    ev["recipients"] = [_rec("甲公司")]
    ev["pipeline"] = [lp_v3.normalize_item({"kind": "attach",
                                            "params": {"files": ["C:/x.pdf"],
                                                       "encrypt": True}})]
    payload = lp_v3.event_export_payload(ev)
    assert payload["format"] == "doc-mailer.event"
    ok, back, err = lp_v3.event_import_payload(payload)
    assert ok and not err
    assert back["name"] == "备份事件"
    assert back["recipients"][0]["name"] == "甲公司"
    assert back["pipeline"][0]["kind"] == "attach"
    assert back["pipeline"][0]["params"]["encrypt"] is True


def test_event_import_rejects_garbage():
    ok, _ev, err = lp_v3.event_import_payload({"hello": "world"})
    assert ok is False
    assert err
    ok, _ev, err = lp_v3.event_import_payload("not a dict")
    assert ok is False


# ---------------------------------------------------------------- 发送模式

def test_send_mode_default_is_per_group():
    assert lp_v3.send_mode_of({}) == "per_group"


def test_send_mode_reads_option():
    assert lp_v3.send_mode_of({"options": {"send_mode": "per_lp"}}) == "per_lp"
    assert lp_v3.send_mode_of({"options": {"send_mode": "all"}}) == "all"


def test_send_mode_falls_back_to_legacy_merge_flag():
    assert lp_v3.send_mode_of({"options": {"merge_groups": False}}) == "per_lp"
    assert lp_v3.send_mode_of({"options": {"merge_groups": True}}) == "per_group"


def test_normalize_options_keeps_send_mode_and_merge_in_sync():
    opts = lp_v3.normalize_options({"send_mode": "per_lp"})
    assert opts["send_mode"] == "per_lp"
    assert opts["merge_groups"] is False
    opts = lp_v3.normalize_options({"send_mode": "per_group"})
    assert opts["merge_groups"] is True


def test_draft_units_per_lp():
    ev = {"options": {"send_mode": "per_lp"},
          "recipients": [_rec("甲"), _rec("乙", group="G1")]}
    units = lp_v3.draft_units(ev)
    assert len(units) == 2


def test_draft_units_per_group():
    ev = {"options": {"send_mode": "per_group"},
          "recipients": [_rec("甲", group="G1"), _rec("乙", group="G1"), _rec("丙")]}
    units = lp_v3.draft_units(ev)
    assert len(units) == 2
    big = max(units, key=lambda u: len(u[1]["recipients"]))
    assert len(big[1]["recipients"]) == 2


def test_draft_units_all_merged():
    ev = {"options": {"send_mode": "all"},
          "recipients": [_rec("甲"), _rec("乙", group="G1")]}
    units = lp_v3.draft_units(ev)
    assert len(units) == 1
    assert len(units[0][1]["recipients"]) == 2


# ---------------------------------------------------------------- 流程图

def test_flow_nodes_structure():
    ev = lp_v3.new_event("流程事件")
    ev["options"]["send_mode"] = "per_lp"
    ev["recipients"] = [_rec("甲"), _rec("乙")]
    ev["pipeline"] = [lp_v3.normalize_item({"kind": "generate",
                                            "params": {"template_path": "C:/t.docx"}}),
                      lp_v3.normalize_item({"kind": "attach",
                                            "params": {"files": ["C:/x.pdf"],
                                                       "encrypt": True}})]
    nodes = lp_v3.flow_nodes(ev)
    kinds = [n["kind"] for n in nodes]
    assert kinds == ["source", "step", "step", "output", "draft"]
    assert nodes[0]["sub"].startswith("2 位")
    assert "加密" in nodes[2]["sub"]
    assert nodes[-1]["title"] == "Outlook 草稿"
    assert "2 封" in nodes[-1]["sub"]


def test_flow_nodes_marks_disabled_step():
    ev = lp_v3.new_event("流程事件")
    st = lp_v3.normalize_item({"kind": "attach",
                               "params": {"files": ["C:/x.pdf"],
                                          "encrypt": True}})
    st["enabled"] = False
    ev["pipeline"] = [st]
    nodes = lp_v3.flow_nodes(ev)
    step = [n for n in nodes if n["kind"] == "step"][0]
    assert step["enabled"] is False
    out = [n for n in nodes if n["kind"] == "output"][0]
    assert out["enabled"] is False       # 没有可用步骤就没有产物
