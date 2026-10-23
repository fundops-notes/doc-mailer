# -*- coding: utf-8 -*-
"""P4 界面测试：收件人库标签页 / 执行历史标签页 / 工作台与收件人库互操作。

设计约束（来自用户 2026-10-23 的反馈与设计稿 §七）：
  * 收件人库与执行历史是**事件库里的标签页**，不占顶层导航；
  * 合并收件人时**绝不覆盖目标已有的打开口令 / 权限口令**（铁律）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

tk = pytest.importorskip("tkinter")

import lp_v3
import lp_tool
from lp_tool import App


@pytest.fixture()
def app(tmp_path, monkeypatch, tk_root):
    monkeypatch.setattr(lp_v3, "default_base_dir", lambda: tmp_path)
    a = App(tk_root)
    yield a
    a.shutdown()


@pytest.fixture(autouse=True)
def no_messagebox(monkeypatch):
    """测试里不弹模态框（模态框会挂住测试进程）。"""
    calls = []

    class _MB:
        @staticmethod
        def showinfo(*a, **_k):
            calls.append(("info", a[0] if a else ""))

        @staticmethod
        def showwarning(*a, **_k):
            calls.append(("warning", a[0] if a else ""))

        @staticmethod
        def showerror(*a, **_k):
            calls.append(("error", a[0] if a else ""))

        @staticmethod
        def askyesno(*a, **_k):
            calls.append(("ask", a[0] if a else ""))
            return True

        @staticmethod
        def askokcancel(*a, **_k):
            calls.append(("ask", a[0] if a else ""))
            return True

    monkeypatch.setattr(lp_tool, "messagebox", _MB)
    return calls


def _pool_rec(app, name, **kw):
    rec = lp_v3.new_pool_recipient({
        "name": name, "wm": name, "to": "%s@example.com" % name,
        "cc": "", "bcc": "", "open_pw": "op-%s" % name,
        "owner_pw": "ow-%s" % name, "fields": {}, "attachments": [],
        "group": ""})
    rec.update(kw)
    ok, err, saved = app.store.save_pool_recipient(rec)
    assert ok, err
    return saved


def _event(app, name, recips=None):
    ev = lp_v3.new_event(name)
    ev["recipients"] = [lp_v3.normalize_recipient(r) for r in (recips or [])]
    app.store.save_event(ev)
    return ev


def _pick_event_stub(app, eid):
    """把 _pick_event 打桩成直接返回指定事件 id（避免测试里弹 Toplevel）。"""
    app._pick_event = lambda _title: eid


def _new_toplevel(app):
    """取 app.root 下面最新的 Toplevel（对话框）。"""
    wins = [w for w in app.root.winfo_children()
            if isinstance(w, tk.Toplevel)]
    assert wins, "没有弹出对话框"
    return wins[-1]


def _walk(w):
    yield w
    for c in w.winfo_children():
        for x in _walk(c):
            yield x


def _click_dialog_button(win, text):
    for w in _walk(win):
        if isinstance(w, tk.Button) or isinstance(w, tk.ttk.Button):
            try:
                if w.cget("text") == text:
                    w.invoke()
                    return True
            except Exception:
                pass
    raise AssertionError("对话框里找不到按钮「%s」" % text)


def _select_in_dialog_tree(win, name=None, index=0):
    """在对话框的表格里选中一行（按第一列的名字选，或按序号）。"""
    for w in _walk(win):
        if isinstance(w, tk.ttk.Treeview):
            rows = w.get_children()
            if name is not None:
                for r in rows:
                    if w.item(r, "values")[0] == name:
                        w.selection_set(r)
                        return r
                raise AssertionError("对话框表格里没有「%s」" % name)
            w.selection_set(rows[index])
            return rows[index]
    raise AssertionError("对话框里找不到表格")


# ------------------------------------------------------------------ 收件人库页

def test_pool_tab_lists_saved_recipients(app):
    _pool_rec(app, "甲公司")
    _pool_rec(app, "乙公司")
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    rows = tree.get_children()
    assert len(rows) == 2
    names = sorted(tree.item(r)["values"][0] for r in rows)
    assert names == ["乙公司", "甲公司"]
    cols = tree["columns"]
    assert "atts" in cols, "收件人库必须有附件路径列（P4 重点）"
    assert "open_pw" in cols and "owner_pw" in cols


def test_pool_tab_shows_attachment_path_and_marks_missing(app, tmp_path):
    real = tmp_path / "季度对账单.pdf"
    real.write_bytes(b"%PDF-1.4 x")
    _pool_rec(app, "甲公司", attachments=[str(real), str(tmp_path / "没了.pdf")])
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    vals = tree.item(tree.get_children()[0])["values"]
    atts = vals[list(tree["columns"]).index("atts")]
    assert "季度对账单.pdf" in atts
    assert "⚠" in atts and "没了.pdf" in atts, "不存在的附件要标 ⚠"


def test_pool_tab_masks_owner_password(app):
    _pool_rec(app, "甲公司")
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    vals = tree.item(tree.get_children()[0])["values"]
    ow = vals[list(tree["columns"]).index("owner_pw")]
    assert ow == "••••", "权限口令在列表里应当打码，实际=%r" % ow
    op = vals[list(tree["columns"]).index("open_pw")]
    assert op == "op-甲公司"


def test_pool_field_columns_collected_from_all_records(app):
    _pool_rec(app, "甲公司", fields={"Call_Amount": "200,000"})
    _pool_rec(app, "乙公司", fields={"Call_Amount": "1", "Due_Date": "2026-11-15"})
    assert app._pool_field_cols() == ["Call_Amount", "Due_Date"]


def test_pool_from_event_pulls_recipients_into_pool(app):
    ev = _event(app, "来源事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op1", "owner_pw": "ow1"}])
    _pick_event_stub(app, ev["id"])
    app.show_screen("home")
    app._pool_from_event()
    pool = app.store.list_pool()
    assert len(pool) == 1
    assert pool[0]["name"] == "甲公司"
    assert pool[0]["open_pw"] == "op1"
    assert pool[0]["owner_pw"] == "ow1"


def test_pool_from_event_never_overwrites_pool_passwords(app):
    """铁律：库里已有的口令，在「从事件导入」时不能被事件里的口令覆盖。"""
    _pool_rec(app, "甲公司", open_pw="op-lib", owner_pw="ow-lib")
    ev = _event(app, "来源事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op-evt", "owner_pw": "ow-evt"}])
    _pick_event_stub(app, ev["id"])
    app.show_screen("home")
    app._pool_from_event()
    pool = app.store.list_pool()
    assert len(pool) == 1
    assert pool[0]["open_pw"] == "op-lib"
    assert pool[0]["owner_pw"] == "ow-lib"


def test_pool_to_event_never_overwrites_event_passwords(app):
    """铁律：事件里已有的打开/权限口令，合并收件人库时绝不被覆盖。"""
    ev = _event(app, "目标事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op-evt", "owner_pw": "ow-evt"}])
    _pool_rec(app, "甲公司", open_pw="op-lib", owner_pw="ow-lib",
              fields={"Call_Amount": "200,000"})
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    tree.selection_set(tree.get_children()[0])
    _pick_event_stub(app, ev["id"])
    app._pool_to_event()
    again, err = app.store.load_event(ev["id"])
    assert err is None
    rec = again["recipients"][0]
    assert rec["open_pw"] == "op-evt"
    assert rec["owner_pw"] == "ow-evt"
    # 空缺仍然会被补上
    assert rec["fields"].get("Call_Amount") == "200,000"


def test_pool_to_event_adds_new_recipient(app):
    ev = _event(app, "目标事件")
    _pool_rec(app, "甲公司", open_pw="op-lib", owner_pw="ow-lib")
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    tree.selection_set(tree.get_children()[0])
    _pick_event_stub(app, ev["id"])
    app._pool_to_event()
    again, _err = app.store.load_event(ev["id"])
    assert len(again["recipients"]) == 1
    assert again["recipients"][0]["name"] == "甲公司"
    assert again["recipients"][0]["open_pw"] == "op-lib"


def test_pool_delete_only_touches_pool(app):
    ev = _event(app, "目标事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op1", "owner_pw": "ow1"}])
    rec = _pool_rec(app, "甲公司")
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    tree.selection_set(tree.get_children()[0])
    app._pool_delete()
    assert app.store.list_pool() == []
    again, _err = app.store.load_event(ev["id"])
    assert len(again["recipients"]) == 1, "删库不影响事件里的收件人"


# ------------------------------------------------------------------ 执行历史页

def _run(app, run_id, started_at, event_id="evt-1", event_name="事件",
         status="ok", items=None, stats=None):
    run = {"id": run_id, "event_id": event_id, "event_name": event_name,
           "started_at": started_at, "finished_at": started_at,
           "status": status, "items": items or [], "logs": ["log line"],
           "errors": [], "warnings": [], "stats": stats or {}}
    ok, err = app.store.save_run(run)
    assert ok, err
    return run


def test_runs_tab_lists_runs_newest_first(app):
    _run(app, "r-old", "2026-10-01 09:00:00")
    _run(app, "r-new", "2026-10-23 09:00:00")
    app.show_screen("home")
    app._refresh_runs_tab()
    tree = app._runs_tree
    rows = tree.get_children()
    assert len(rows) == 2
    assert tree.item(rows[0])["values"][0] == "2026-10-23 09:00:00"
    assert tree.item(rows[1])["values"][0] == "2026-10-01 09:00:00"


def test_runs_tab_shows_draft_and_file_counts(app):
    _run(app, "r1", "2026-10-23 09:00:00", items=[
        {"name": "甲公司", "members": ["a@x.com"], "files": ["a.pdf", "b.pdf"],
         "missing": [], "draft": "draft-1", "error": ""},
        {"name": "乙公司", "members": ["b@x.com"], "files": [],
         "missing": ["missing.pdf"], "draft": "", "error": "生成失败"}],
        stats={"ok": 1, "fail": 1, "skip": 0})
    app.show_screen("home")
    app._refresh_runs_tab()
    tree = app._runs_tree
    vals = tree.item(tree.get_children()[0])["values"]
    cols = list(tree["columns"])
    assert vals[cols.index("drafts")] == "1/2"
    assert vals[cols.index("ok")] == 1
    assert vals[cols.index("fail")] == 1
    assert vals[cols.index("files")] == 2
    assert vals[cols.index("status")] == "已完成"


def test_runs_tab_labels_missing_event(app):
    _run(app, "r1", "2026-10-23 09:00:00", event_name="")
    app.show_screen("home")
    app._refresh_runs_tab()
    tree = app._runs_tree
    vals = tree.item(tree.get_children()[0])["values"]
    assert vals[1] == "（事件已删除）"


def test_runs_open_event_navigates_to_workbench(app):
    ev = _event(app, "目标事件")
    _run(app, "r1", "2026-10-23 09:00:00", event_id=ev["id"],
         event_name="目标事件")
    app.show_screen("home")
    app._refresh_runs_tab()
    tree = app._runs_tree
    tree.selection_set(tree.get_children()[0])
    app._runs_open_event()
    assert app._wb_event is not None
    assert app._wb_event["id"] == ev["id"]


# ------------------------------------------------------- 工作台 ↔ 收件人库互操作

def test_workbench_save_to_pool_stores_event_recipients(app):
    ev = _event(app, "来源事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op1", "owner_pw": "ow1"}])
    app.open_event(ev["id"])
    app._wb_save_to_pool()
    pool = app.store.list_pool()
    assert len(pool) == 1
    assert pool[0]["name"] == "甲公司"
    assert pool[0]["open_pw"] == "op1"
    assert pool[0]["owner_pw"] == "ow1"


def test_workbench_save_to_pool_keeps_existing_pool_passwords(app):
    """铁律：「存入收件人库」时库里已有的口令保留。"""
    _pool_rec(app, "甲公司", open_pw="op-lib", owner_pw="ow-lib")
    ev = _event(app, "来源事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op-evt", "owner_pw": "ow-evt"}])
    app.open_event(ev["id"])
    app._wb_save_to_pool()
    pool = app.store.list_pool()
    assert len(pool) == 1
    assert pool[0]["open_pw"] == "op-lib"
    assert pool[0]["owner_pw"] == "ow-lib"


def test_workbench_add_from_pool_merges_into_event(app):
    ev = _event(app, "目标事件", [{"name": "甲公司", "to": "a@x.com",
                                   "open_pw": "op-evt", "owner_pw": "ow-evt"}])
    _pool_rec(app, "甲公司", open_pw="op-lib", owner_pw="ow-lib")
    _pool_rec(app, "乙公司", open_pw="op2", owner_pw="ow2")
    app.open_event(ev["id"])
    app._wb_add_from_pool()
    win = _new_toplevel(app)
    _select_in_dialog_tree(win, name="乙公司")
    _click_dialog_button(win, "添加")
    again, _err = app.store.load_event(ev["id"])
    names = [r["name"] for r in again["recipients"]]
    assert names == ["甲公司", "乙公司"]
    assert again["recipients"][0]["open_pw"] == "op-evt", "已有口令不能被覆盖"
    assert again["recipients"][0]["owner_pw"] == "ow-evt"
    assert again["recipients"][1]["open_pw"] == "op2"
    assert again["recipients"][1]["owner_pw"] == "ow2"


def test_workbench_recipients_step_has_pool_buttons(app):
    """事件工作台「收件人」步骤工具条上有「从收件人库添加…」与「存入收件人库」。"""
    ev = _event(app, "目标事件")
    app.open_event(ev["id"])
    app._wb_show_step(0)
    labels = []

    def walk(w):
        for c in w.winfo_children():
            try:
                if c.cget("text"):
                    labels.append(str(c.cget("text")))
            except Exception:
                pass
            walk(c)

    walk(app._wb_content)
    assert any("从收件人库添加" in t for t in labels), labels
    assert any("存入收件人库" in t for t in labels), labels
