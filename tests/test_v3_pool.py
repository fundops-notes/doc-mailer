# -*- coding: utf-8 -*-
"""P4 收尾：收件人库（跨事件复用）+ 执行历史的数据层测试。

重点是那条铁律：**合并时绝不覆盖目标已有的打开/权限口令**。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

import lp_v3


def _rec(name="甲公司", **kw):
    r = {"name": name, "to": "a@x.com", "open_pw": "open123", "owner_pw": "owner123"}
    r.update(kw)
    return lp_v3.normalize_recipient(r)


# ---------------------------------------------------------------- 规范化

def test_normalize_pool_recipient_keeps_id():
    r = lp_v3.normalize_pool_recipient({"name": "甲公司", "id": "rcp_1"})
    assert r["id"] == "rcp_1"
    assert r["name"] == "甲公司"
    assert r["attachments"] == []


def test_normalize_pool_recipient_generates_id():
    r = lp_v3.normalize_pool_recipient({"name": "甲公司"})
    assert r["id"].startswith("rcp-")


def test_new_pool_recipient_has_id():
    r = lp_v3.new_pool_recipient({"name": "甲公司"})
    assert r["id"].startswith("rcp-")
    assert r["updated_at"]


# ---------------------------------------------------------------- 合并（铁律）

def test_merge_into_list_adds_new():
    lst = [_rec("甲公司")]
    lst, added, updated = lp_v3.merge_recipient_into_list(lst, _rec("乙公司"))
    assert added == 1 and updated == 0
    assert len(lst) == 2


def test_merge_into_list_never_overwrites_passwords():
    """库里那份是旧口令 → 不能覆盖事件里已用的口令。"""
    lst = [_rec("甲公司", open_pw="NEW-open", owner_pw="NEW-owner")]
    lst, added, updated = lp_v3.merge_recipient_into_list(
        lst, _rec("甲公司", open_pw="OLD-open", owner_pw="OLD-owner"))
    assert added == 0 and updated == 1
    assert lst[0]["open_pw"] == "NEW-open"
    assert lst[0]["owner_pw"] == "NEW-owner"


def test_merge_into_list_fills_empty_only():
    lst = [_rec("甲公司", to="", cc="cc@x.com")]
    lst, _a, _u = lp_v3.merge_recipient_into_list(
        lst, _rec("甲公司", to="new@x.com", cc="other@x.com"))
    assert lst[0]["to"] == "new@x.com"       # 空 → 补上
    assert lst[0]["cc"] == "cc@x.com"        # 已有 → 不动


def test_merge_into_list_keeps_existing_custom_fields():
    lst = [_rec("甲公司", fields={"Call_Amount": "200,000"})]
    lst, _a, _u = lp_v3.merge_recipient_into_list(
        lst, _rec("甲公司", fields={"Call_Amount": "999", "Extra": "x"}))
    assert lst[0]["fields"]["Call_Amount"] == "200,000"
    assert lst[0]["fields"]["Extra"] == "x"


def test_merge_into_list_merges_attachment_paths():
    lst = [_rec("甲公司", attachments=[r"C:\a.pdf"])]
    lst, _a, _u = lp_v3.merge_recipient_into_list(
        lst, _rec("甲公司", attachments=[r"C:\b.pdf"]))
    assert lst[0]["attachments"] == [r"C:\a.pdf", r"C:\b.pdf"]


def test_merge_into_list_ignores_empty_name():
    lst = [_rec("甲公司")]
    lst, added, updated = lp_v3.merge_recipient_into_list(lst, {"name": ""})
    assert added == 0 and updated == 0
    assert len(lst) == 1


# ---------------------------------------------------------------- 收件人库

def test_pool_merge_adds_new():
    pool = []
    pool, new_rec, is_update = lp_v3.pool_merge_recipient(pool, _rec("甲公司"))
    assert new_rec is not None and is_update is False
    assert len(pool) == 1


def test_pool_merge_updates_existing_without_touching_passwords():
    pool = [lp_v3.normalize_pool_recipient(
        _rec("甲公司", open_pw="KEEP-open", owner_pw="KEEP-owner"))]
    pool, new_rec, is_update = lp_v3.pool_merge_recipient(
        pool, _rec("甲公司", open_pw="OLD-open", owner_pw="OLD-owner", to=""))
    assert new_rec is None and is_update is True
    assert pool[0]["open_pw"] == "KEEP-open"
    assert pool[0]["owner_pw"] == "KEEP-owner"


# ---------------------------------------------------------------- 并入事件

def test_merge_recipients_into_event_keeps_event_passwords():
    ev = lp_v3.new_event("事件一")
    ev["recipients"] = [_rec("甲公司", open_pw="EVENT-open", owner_pw="EVENT-owner")]
    new_ev, added, updated = lp_v3.merge_recipients_into_event(
        ev, [_rec("甲公司", open_pw="POOL-open", owner_pw="POOL-owner"),
             _rec("乙公司")])
    assert added == 1 and updated == 1
    keep = [r for r in new_ev["recipients"] if r["name"] == "甲公司"][0]
    assert keep["open_pw"] == "EVENT-open"
    assert keep["owner_pw"] == "EVENT-owner"
    assert len(new_ev["recipients"]) == 2


# ---------------------------------------------------------------- Store 存取

def test_store_pool_roundtrip(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    rec = lp_v3.new_pool_recipient(_rec("甲公司"))
    ok, err, saved = store.save_pool_recipient(rec)
    assert ok, err
    pool = store.list_pool()
    assert len(pool) == 1
    assert pool[0]["name"] == "甲公司"
    assert pool[0]["id"] == saved["id"]


def test_store_pool_sorted_by_name(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    for n in ("丙公司", "甲公司", "乙公司"):
        store.save_pool_recipient(lp_v3.new_pool_recipient(_rec(n)))
    assert [r["name"] for r in store.list_pool()] == ["乙公司", "丙公司", "甲公司"] \
        or [r["name"] for r in store.list_pool()] == sorted(
            ["甲公司", "乙公司", "丙公司"])


def test_store_pool_delete(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    rec = lp_v3.new_pool_recipient(_rec("甲公司"))
    store.save_pool_recipient(rec)
    ok, err = store.delete_pool_recipient(rec["id"])
    assert ok, err
    assert store.list_pool() == []


def test_store_pool_corrupt_file_is_quarantined(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    store.ensure_dirs()
    bad = store.recipients_dir / "bad.json"
    bad.write_text("{ 这不是 json", encoding="utf-8")
    assert store.list_pool() == []
    # 损坏文件被改名留证，不删除
    left = list(store.recipients_dir.glob("*.json"))
    assert all("corrupt" in p.name for p in left)


def test_store_ensure_dirs_creates_recipients_dir(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    store.ensure_dirs()
    assert store.recipients_dir.exists()


# ---------------------------------------------------------------- 执行历史

def test_list_runs_sorted_newest_first(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    store.ensure_dirs()
    for started in ("2026-10-01 10:00:00", "2026-10-03 10:00:00"):
        store.save_run({"id": "run_%s" % started, "event_id": "e1",
                        "event_name": "事件", "started_at": started,
                        "items": [], "stats": {"ok": 0, "fail": 0, "skip": 0}})
    runs = store.list_runs()
    assert len(runs) == 2
    assert runs[0]["started_at"] > runs[1]["started_at"]


def test_list_runs_filters_by_event(tmp_path):
    store = lp_v3.Store(base_dir=tmp_path)
    store.ensure_dirs()
    store.save_run({"id": "r1", "event_id": "e1", "started_at": "2026-10-01 10:00:00"})
    store.save_run({"id": "r2", "event_id": "e2", "started_at": "2026-10-02 10:00:00"})
    assert len(store.list_runs("e1")) == 1
    assert store.list_runs("e1")[0]["event_id"] == "e1"
    assert len(store.list_runs()) == 2

# ---------------------------------------------------------------- 示例数据播种

def test_seed_demo_seeds_pool(tmp_path):
    """首次启动要播种一份示例收件人库（让「收件人库」标签页不是空的）。"""
    ex = tmp_path / "lp_data.example.json"
    ex.write_text(
        '{"lps": [{"name": "甲公司", "to": "a@x.com", "open_pw": "", '
        '"owner_pw": ""}], "pool": [{"name": "乙公司", "to": "b@x.com"}]}',
        encoding="utf-8")
    store = lp_v3.Store(base_dir=tmp_path)
    ok, notes = lp_v3.seed_demo(store, ex)
    assert ok, notes
    pool = store.list_pool()
    assert len(pool) == 1
    assert pool[0]["name"] == "乙公司"
    assert pool[0]["open_pw"], "示例库记录也要有打开口令"
    assert pool[0]["owner_pw"], "示例库记录也要有权限口令"


def test_seed_demo_has_no_scenario_and_empty_pipeline(tmp_path):
    """v3.5.1：场景模板移除后，示例数据只有事件，且流水线为空（不挡首次运行）。"""
    ex = tmp_path / "lp_data.example.json"
    ex.write_text(
        '{"lps": [{"name": "甲公司", "to": "a@x.com", "open_pw": "o1", '
        '"owner_pw": "w1"}]}', encoding="utf-8")
    store = lp_v3.Store(base_dir=tmp_path)
    lp_v3.seed_demo(store, ex)
    assert store.list_scenarios() == []
    ev = store.list_events()[0]
    assert ev["pipeline"] == []
    assert ev["recipients"][0]["name"] == "甲公司"
