# -*- coding: utf-8 -*-
"""v3 数据层回归测试（目录化存储 / 数据模型 / v2 迁移）：pytest tests/"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

from lp_v3 import (DEFAULT_FILENAME_TPL, Store, migrate_v2_to_v3,
                   needs_migration, new_event, new_scenario, normalize_attachment,
                   normalize_event, normalize_item, normalize_recipient,
                   read_json_file, split_paths, sync_scenario_to_event,
                   v2_to_v3_payload, write_json_file)


# ---------------------------------------------------------------- 数据模型

def test_normalize_recipient_accepts_v2_shape():
    """v2 的 files / merge_group / wm 要能直接吃进 v3 模型。"""
    r = normalize_recipient({
        "name": "甲公司", "wm": "甲水印", "to": "a@x.com",
        "files": "C:/a.pdf;D:/b,c.pdf", "merge_group": "G1",
        "open_pw": "op", "owner_pw": "ow",
    })
    assert r["name"] == "甲公司" and r["wm"] == "甲水印"
    assert r["group"] == "G1"
    assert r["attachments"] == ["C:/a.pdf", "D:/b,c.pdf"]
    assert r["fields"] == {}


def test_normalize_recipient_keeps_custom_fields():
    r = normalize_recipient({"name": "甲", "fields": {"Call_Amount": "200,000"}})
    assert r["fields"]["Call_Amount"] == "200,000"


def test_normalize_item_rejects_unknown_kind():
    """v3.1：旧的六种步骤类型不再出现在模型里。"""
    s = normalize_item({"kind": "不存在的类型"})
    assert s["kind"] == "generate"
    assert s["granularity"] == "per_group"
    assert s["id"]


def test_normalize_attachment_scope():
    a = normalize_attachment({"path": "x.pdf", "scope": "group", "group": "G"})
    assert a["scope"] == "group" and a["group"] == "G"
    b = normalize_attachment({"path": "x.pdf", "scope": "乱写"})
    assert b["scope"] == "all"


def test_event_from_scenario_is_a_copy():
    """复制式：事件改配置不影响场景。"""
    sc = new_scenario("场景A")
    sc["recipients"].append(normalize_recipient({"name": "甲", "to": "a@x.com"}))
    ev = new_event("事件A", sc)
    ev["recipients"].append(normalize_recipient({"name": "乙", "to": "b@x.com"}))
    ev["options"]["outdir"] = "D:/out"
    assert len(sc["recipients"]) == 1, "改事件不应改场景"
    assert sc["options"].get("outdir", "") != "D:/out"
    assert ev["scenario_id"] == sc["id"]


def test_sync_scenario_to_event_keeps_passwords():
    """手动同步：只补空缺，已有口令与自定义字段必须保住。"""
    sc = new_scenario("场景A")
    sc["recipients"].append(normalize_recipient({"name": "甲", "to": "a@x.com"}))
    ev = new_event("事件A")
    ev["recipients"].append(normalize_recipient(
        {"name": "甲", "to": "a@x.com", "open_pw": "keepme", "owner_pw": "keepowner",
         "fields": {"X": "1"}}))
    ev2, notes = sync_scenario_to_event(sc, ev)
    r = ev2["recipients"][0]
    assert r["open_pw"] == "keepme" and r["owner_pw"] == "keepowner"
    assert r["fields"]["X"] == "1"
    assert any("收件人" in n for n in notes)


# ---------------------------------------------------------------- 存储

def test_store_roundtrip(tmp_path):
    s = Store(tmp_path)
    s.ensure_dirs()
    ev = new_event("事件1")
    ev["recipients"].append(normalize_recipient({"name": "甲", "to": "a@x.com"}))
    ok, err = s.save_event(ev)
    assert ok and not err
    got, err = s.load_event(ev["id"])
    assert err is None
    assert got["name"] == "事件1"
    assert got["recipients"][0]["name"] == "甲"
    # 场景
    sc = new_scenario("场景1")
    ok, err = s.save_scenario(sc)
    assert ok and not err
    assert [x["name"] for x in s.list_scenarios()] == ["场景1"]
    # 设置
    st = {"version": 3, "outdir": "D:/x"}
    ok, err = s.save_settings(st)
    assert ok and not err
    got_st, err = s.load_settings()
    assert got_st["outdir"] == "D:/x"


def test_store_lists_events_sorted_by_updated(tmp_path):
    s = Store(tmp_path)
    s.ensure_dirs()
    # 直接写盘：save_event 会把 updated_at 盖成当前时间，排序要靠固定值才稳定
    for name, ts in (("A", "2020-01-01 00:00:00"), ("B", "2099-01-01 00:00:00")):
        ev = new_event(name)
        ev["updated_at"] = ts
        s.event_path(ev["id"]).write_text(
            json.dumps(ev, ensure_ascii=False), encoding="utf-8")
    names = [e["name"] for e in s.list_events()]
    assert names[0] == "B"


def test_corrupt_event_file_is_quarantined_not_overwritten(tmp_path):
    """损坏数据绝不覆盖：坏文件改名留证，不被新数据顶掉。"""
    s = Store(tmp_path)
    s.ensure_dirs()
    p = s.event_path("evt-bad")
    p.write_text("{ 这不是合法 JSON", encoding="utf-8")
    assert s.list_events() == []
    assert not p.exists()
    quarantined = list(s.events_dir.glob("event.corrupt-*.json"))
    assert len(quarantined) == 1
    assert "不是合法 JSON" in quarantined[0].read_text(encoding="utf-8")


def test_write_json_file_keeps_bak(tmp_path):
    p = tmp_path / "x.json"
    assert write_json_file(p, {"a": 1})[0]
    assert write_json_file(p, {"a": 2})[0]
    assert json.loads(p.read_text(encoding="utf-8"))["a"] == 2
    assert json.loads((tmp_path / "x.json.bak").read_text(encoding="utf-8"))["a"] == 1


# ---------------------------------------------------------------- 迁移

def _legacy_file(tmp_path, payload):
    p = tmp_path / "lp_data.json"
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return p


def test_migration_creates_directory_and_keeps_original(tmp_path):
    legacy = _legacy_file(tmp_path, {
        "lps": [{"name": "甲公司", "to": "a@x.com", "open_pw": "op1", "owner_pw": "ow1",
                 "files": "C:/a.pdf", "merge_group": "G1"}],
        "subject": "标题 {{收件人名称}}", "body": "<p>正文</p>",
        "wm_settings": {"opacity": 0.3}, "gp_owner_mode": True, "gp_owner_pw": "gp",
        "outdir": "D:/out", "filename_tpl": DEFAULT_FILENAME_TPL,
        "word_template": "C:/tpl.docx",
    })
    s = Store(tmp_path)
    assert needs_migration(s)
    ok, notes, err = migrate_v2_to_v3(s)
    assert ok, err
    assert (tmp_path / "lp_data" / "settings.json").exists()
    assert list((tmp_path / "lp_data" / "events").glob("*.json"))
    assert list((tmp_path / "lp_data" / "scenarios").glob("*.json"))
    # 原文件改名保留，不删除
    assert not legacy.exists()
    assert (tmp_path / "lp_data.json.bak").exists()
    # 内容对得上
    st, _ = s.load_settings()
    assert st["wm_settings"]["opacity"] == 0.3 and st["gp_owner_mode"] is True
    evs = s.list_events()
    assert len(evs) == 1
    r = evs[0]["recipients"][0]
    assert r["name"] == "甲公司" and r["group"] == "G1"
    assert r["attachments"] == ["C:/a.pdf"]
    assert r["open_pw"] == "op1" and r["owner_pw"] == "ow1"
    # 流水线：Word 模板 → 一个「生成文档」条目（带转 PDF）
    assert len(evs[0]["pipeline"]) == 1
    it = evs[0]["pipeline"][0]
    assert it["kind"] == "generate"
    assert it["params"]["template_path"] == "C:/tpl.docx"
    assert it["params"]["to_pdf"] is True
    assert evs[0]["mail"]["subject"] == "标题 {{收件人名称}}"


def test_migration_skipped_when_data_dir_exists(tmp_path):
    s = Store(tmp_path)
    s.ensure_dirs()
    s.save_settings({"version": 3})
    _legacy_file(tmp_path, {"lps": []})
    ok, notes, err = migrate_v2_to_v3(s)
    assert not ok and "跳过迁移" in err


def test_migration_refuses_to_overwrite_corrupt_legacy(tmp_path):
    """损坏的 lp_data.json 不被迁移覆盖：改名留证后中止。"""
    legacy = tmp_path / "lp_data.json"
    legacy.write_text("{坏掉的 JSON", encoding="utf-8")
    s = Store(tmp_path)
    ok, notes, err = migrate_v2_to_v3(s)
    assert not ok
    assert err and "读取失败" in err
    assert not legacy.exists()
    assert list(tmp_path.glob("lp_data.corrupt-*.json"))
    assert not (tmp_path / "lp_data" / "settings.json").exists()


def test_seed_demo_creates_scenario_and_event(tmp_path):
    """首次启动播种示例数据：v2 的「开箱有演示数据」体验要保留。"""
    from lp_v3 import seed_demo
    s = Store(tmp_path)
    ok, notes = seed_demo(s)
    assert ok
    assert len(s.list_scenarios()) == 1
    evs = s.list_events()
    assert len(evs) == 1
    recips = evs[0]["recipients"]
    assert recips
    # 口令必须补齐，且打开口令与权限口令不同
    for r in recips:
        assert r["open_pw"] and r["owner_pw"]
        assert r["open_pw"] != r["owner_pw"]
    # 打开口令互不相同（否则一个人能开别人的文件）
    assert len(set(r["open_pw"] for r in recips)) == len(recips)


def test_seed_demo_uses_example_file(tmp_path):
    from lp_v3 import seed_demo
    ex = tmp_path / "lp_data.example.json"
    ex.write_text(json.dumps({
        "lps": [{"name": "甲公司", "to": "a@demo.cn", "open_pw": "", "owner_pw": ""}],
        "subject": "S", "body": "<p>B</p>"}, ensure_ascii=False), encoding="utf-8")
    s = Store(tmp_path)
    seed_demo(s)
    evs = s.list_events()
    assert evs[0]["recipients"][0]["name"] == "甲公司"
    assert evs[0]["mail"]["subject"] == "S"


def test_v2_payload_without_word_template_notes_pdf_gap():
    _settings, _sc, ev, notes = v2_to_v3_payload({"lps": [], "subject": "s"})
    assert ev["pipeline"] == []
    assert any("源 PDF" in n for n in notes)


def test_split_paths_keeps_comma_in_filename():
    assert split_paths("C:/a.pdf;" + chr(10) + "D:/X, L.P..pdf") == [
        "C:/a.pdf", "D:/X, L.P..pdf"]
