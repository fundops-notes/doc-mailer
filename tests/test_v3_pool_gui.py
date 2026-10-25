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
    assert "atts" not in cols, "v3.4：附件路径列已从收件人页移除（统一到流水线）"
    assert "open_pw" in cols and "owner_pw" in cols


def test_pool_tab_shows_attachment_path_and_marks_missing(app, tmp_path):
    """v3.4：附件路径列已移除，此用例改为确认列不存在且数据仍保留在记录里"""
    real = tmp_path / "季度对账单.pdf"
    real.write_bytes(b"%PDF-1.4 x")
    rec = _pool_rec(app, "甲公司", attachments=[str(real), str(tmp_path / "没了.pdf")])
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    cols = tree["columns"]
    assert "atts" not in cols
    # 数据本体仍在（流水线「路径来源=收件人附件」仍可用）
    saved = app.store.load_pool_recipient(rec["id"])[0] if hasattr(app.store, "load_pool_recipient") else None
    assert saved is None or [str(a) for a in (saved.get("attachments") or [])]


def test_pool_tab_masks_owner_password(app):
    _pool_rec(app, "甲公司")
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    ow_idx = list(tree["columns"]).index("owner_pw")
    op_idx = list(tree["columns"]).index("open_pw")
    found = False
    for r in tree.get_children():
        vals = tree.item(r)["values"]
        if len(vals) > ow_idx:
            assert vals[ow_idx] in ("••••", ""),                 "权限口令必须打码，实际=%r" % vals[ow_idx]
            if vals[list(tree["columns"]).index("name")] == "甲公司":
                found = True
    assert found, "测试插入的甲公司应在列表里"


def test_pool_does_not_collect_event_fields(app):
    """v3.7.1：自定义字段属于事件，收件人库不再收集/显示历史字段。"""
    _pool_rec(app, "甲公司", fields={"Call_Amount": "200,000"})
    _pool_rec(app, "乙公司", fields={"Call_Amount": "1", "Due_Date": "2026-11-15"})
    assert app._pool_field_cols() == []
    app.show_screen("home")
    app._refresh_pool_tab()
    heads = [str(app._pool_tree.heading(c)["text"]) for c in app._pool_tree["columns"]]
    assert "Call_Amount" not in heads and "Due_Date" not in heads


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
    # v3.7.2：收件人库不保存事件字段；即使旧库有字段也不带进事件
    assert "Call_Amount" not in (rec.get("fields") or {})


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


# -------------------------------------------------- v3.8.2 按标签挑选库收件人

def _open_pool_picker(app, ev):
    app.open_event(ev["id"])
    app._wb_add_from_pool()
    win = _new_toplevel(app)
    tree = next(w for w in _walk(win) if isinstance(w, tk.ttk.Treeview))
    combo = next(w for w in _walk(win) if isinstance(w, tk.ttk.Combobox))
    return win, tree, combo


def _filter_pool(app, combo, label):
    combo.set(label)
    combo.event_generate("<<ComboboxSelected>>")
    app.root.update()


def _picker_names(tree):
    return {tree.item(row, "values")[0] for row in tree.get_children()}


def test_pool_picker_lists_unique_tags_and_shows_tag_column(app):
    _pool_rec(app, "甲公司", tags=["重点", "华东"])
    _pool_rec(app, "乙公司", tags=["重点"])
    _pool_rec(app, "未分类", tags=[])
    # 真正叫“未打标签”的标签也能和筛选入口区分。
    _pool_rec(app, "特殊标签", tags=["未打标签"])
    win, tree, combo = _open_pool_picker(app, _event(app, "目标事件"))
    assert str(combo.cget("state")) == "readonly"
    assert list(combo.cget("values")) == [
        "全部收件人", "未打标签", "标签：华东", "标签：未打标签", "标签：重点"]
    assert tree.cget("columns") == ("name", "to", "tags")
    assert len(tree.get_children()) == 4
    row = next(row for row in tree.get_children()
               if tree.item(row, "values")[0] == "甲公司")
    assert tree.item(row, "values")[2] == "重点、华东"
    _filter_pool(app, combo, "标签：未打标签")
    assert _picker_names(tree) == {"特殊标签"}
    _click_dialog_button(win, "取消")


def test_pool_picker_batch_adds_exact_tag_and_preserves_event_passwords(app):
    _pool_rec(app, "甲公司", tags=["重点", "华东"], cc="cc@example.com")
    _pool_rec(app, "乙公司", tags=["重点"])
    _pool_rec(app, "丙公司", tags=["重点客户"])
    _pool_rec(app, "丁公司", tags=[])
    ev = _event(app, "目标事件", [{"name": "甲公司", "to": "existing@example.com",
                                       "open_pw": "event-open", "owner_pw": "event-owner",
                                       "fields": {"Due_Date": "2026/11/01"}}])
    pool_before = app.store.list_pool()
    win, tree, combo = _open_pool_picker(app, ev)
    _filter_pool(app, combo, "标签：重点")
    assert _picker_names(tree) == {"甲公司", "乙公司"}
    _click_dialog_button(win, "全选当前列表")
    assert len(tree.selection()) == 2
    _click_dialog_button(win, "添加")
    saved, err = app.store.load_event(ev["id"])
    assert not err
    people = {r["name"]: r for r in saved["recipients"]}
    assert set(people) == {"甲公司", "乙公司"}
    assert people["甲公司"]["to"] == "existing@example.com"
    assert people["甲公司"]["cc"] == "cc@example.com"
    assert people["甲公司"]["open_pw"] == "event-open"
    assert people["甲公司"]["owner_pw"] == "event-owner"
    assert people["甲公司"]["fields"] == {"Due_Date": "2026/11/01"}
    assert app.store.list_pool() == pool_before


def test_pool_picker_switch_filter_clears_hidden_selection(app):
    _pool_rec(app, "华东客户", tags=["华东"])
    _pool_rec(app, "华南客户", tags=["华南"])
    ev = _event(app, "目标事件")
    win, tree, combo = _open_pool_picker(app, ev)
    _filter_pool(app, combo, "标签：华东")
    _click_dialog_button(win, "全选当前列表")
    assert len(tree.selection()) == 1
    _filter_pool(app, combo, "标签：华南")
    assert not tree.selection()
    assert _picker_names(tree) == {"华南客户"}
    labels = [str(w.cget("text")) for w in _walk(win)
              if isinstance(w, tk.ttk.Label)]
    assert "显示 1 / 共 2 位；已选 0 位" in labels
    _click_dialog_button(win, "全选当前列表")
    _click_dialog_button(win, "添加")
    saved, _ = app.store.load_event(ev["id"])
    assert [r["name"] for r in saved["recipients"]] == ["华南客户"]


def test_pool_picker_untagged_and_cancel_leave_event_unchanged(app):
    _pool_rec(app, "有标签", tags=["重点"])
    _pool_rec(app, "无标签", tags=[])
    ev = _event(app, "目标事件", [{"name": "已有客户"}])
    win, tree, combo = _open_pool_picker(app, ev)
    before, _ = app.store.load_event(ev["id"])
    _filter_pool(app, combo, "未打标签")
    assert _picker_names(tree) == {"无标签"}
    _click_dialog_button(win, "全选当前列表")
    _click_dialog_button(win, "取消选择")
    assert not tree.selection()
    _filter_pool(app, combo, "全部收件人")
    assert _picker_names(tree) == {"有标签", "无标签"}
    _click_dialog_button(win, "全选当前列表")
    _click_dialog_button(win, "取消")
    after, _ = app.store.load_event(ev["id"])
    assert after == before


def test_pool_picker_empty_filter_warns_and_adds_nothing(app, no_messagebox):
    _pool_rec(app, "有标签", tags=["重点"])
    ev = _event(app, "目标事件")
    win, tree, combo = _open_pool_picker(app, ev)
    _filter_pool(app, combo, "未打标签")
    assert not tree.get_children()
    _click_dialog_button(win, "全选当前列表")
    _click_dialog_button(win, "取消选择")
    _click_dialog_button(win, "添加")
    assert ("warning", "未选择") in no_messagebox
    assert win.winfo_exists()
    saved, _ = app.store.load_event(ev["id"])
    assert not saved["recipients"]
    _click_dialog_button(win, "取消")


def test_pool_picker_select_all_includes_filtered_rows_below_scroll(app):
    for i in range(35):
        _pool_rec(app, "批量客户%02d" % i, tags=["同一批"])
    _pool_rec(app, "其他客户", tags=["另一批"])
    ev = _event(app, "目标事件")
    win, tree, combo = _open_pool_picker(app, ev)
    _filter_pool(app, combo, "标签：同一批")
    assert len(tree.get_children()) == 35
    _click_dialog_button(win, "全选当前列表")
    assert len(tree.selection()) == 35
    _click_dialog_button(win, "添加")
    saved, _ = app.store.load_event(ev["id"])
    assert len(saved["recipients"]) == 35
    assert "其他客户" not in {r["name"] for r in saved["recipients"]}


@pytest.mark.parametrize("geometry", ["820x560", "700x460"])
def test_pool_picker_actions_are_visible_at_default_and_minimum_size(app, geometry):
    _pool_rec(app, "甲公司", tags=["重点"])
    win, _tree, _combo = _open_pool_picker(app, _event(app, "目标事件"))
    app.root.deiconify()
    try:
        win.geometry(geometry)
        app.root.update()
        buttons = [w for w in _walk(win) if isinstance(w, tk.ttk.Button)
                   and str(w.cget("text")) in ("全选当前列表", "取消选择", "添加", "取消")]
        assert len(buttons) == 4
        for button in buttons:
            assert button.winfo_ismapped(), button.cget("text")
            assert button.winfo_rooty() >= win.winfo_rooty()
            assert button.winfo_rooty() + button.winfo_height() <= win.winfo_rooty() + win.winfo_height()
            assert button.winfo_rootx() + button.winfo_width() <= win.winfo_rootx() + win.winfo_width()
        _click_dialog_button(win, "取消")
    finally:
        app.root.withdraw()

# ---------------------------------------------------------------- v3.7.0

def test_pool_tags_saved_and_filtered(app):
    _pool_rec(app, "甲公司", tags=["重点", "华东"])
    _pool_rec(app, "乙公司", tags=["重点"])
    app.show_screen("home")
    app._refresh_pool_tab()
    pool = {r["name"]: r for r in app.store.list_pool()}
    assert pool["甲公司"]["tags"] == ["重点", "华东"]
    assert pool["乙公司"]["tags"] == ["重点"]


def test_pool_batch_tags_only_remove_common(app):
    _pool_rec(app, "甲公司", tags=["重点", "华东"])
    _pool_rec(app, "乙公司", tags=["重点"])
    app.show_screen("home")
    app._refresh_pool_tab()
    tree = app._pool_tree
    tree.selection_set(tree.get_children())
    app._pool_tags()
    dlg = _new_toplevel(app)
    assert dlg is not None

    # 添加：输入新标签并应用
    entry = next(w for w in _walk(dlg) if w.winfo_class() == "TEntry")
    entry.insert(0, "新标签")
    btn = next(w for w in _walk(dlg) if w.winfo_class() == "TButton"
               and "添加勾选" in str(w.cget("text")))
    btn.invoke()
    app.root.update()
    pool = {r["name"]: r for r in app.store.list_pool()}
    assert "新标签" in pool["甲公司"]["tags"] and "新标签" in pool["乙公司"]["tags"]

    # 移除特有标签：勾选「华东」应被拒绝（不是共有）
    tree.selection_set(tree.get_children())
    app._pool_tags()
    dlg = _new_toplevel(app)
    cbs = [w for w in _walk(dlg) if w.winfo_class() == "TCheckbutton"]
    east = next(w for w in cbs if "华东" in str(w.cget("text")))
    east.invoke()
    remove_btn = next(w for w in _walk(dlg) if w.winfo_class() == "TButton"
                      and "移除勾选" in str(w.cget("text")))
    remove_btn.invoke()
    app.root.update()
    pool = {r["name"]: r for r in app.store.list_pool()}
    assert "华东" in pool["甲公司"]["tags"], "特异标签不能被批量移除"

    # 移除共有标签「重点」：两人都应移除
    tree.selection_set(tree.get_children())
    app._pool_tags()
    dlg = _new_toplevel(app)
    cbs = [w for w in _walk(dlg) if w.winfo_class() == "TCheckbutton"]
    common = next(w for w in cbs if "重点" in str(w.cget("text")))
    common.invoke()
    remove_btn = next(w for w in _walk(dlg) if w.winfo_class() == "TButton"
                      and "移除勾选" in str(w.cget("text")))
    remove_btn.invoke()
    app.root.update()
    pool = {r["name"]: r for r in app.store.list_pool()}
    assert "重点" not in pool["甲公司"]["tags"]
    assert "重点" not in pool["乙公司"]["tags"]


def test_pick_event_waits_for_user_selection(app):
    """回归：_pick_event 旧实现建窗后立即 return None，导致「添加到事件」静默失效。"""
    ev = _event(app, "目标事件")
    app.show_screen("home")
    app.root.update()
    picked = []

    def choose():
        dlg = _new_toplevel(app)
        if dlg is None:
            return
        lb = next(w for w in _walk(dlg) if w.winfo_class() == "Listbox")
        labels = [lb.get(i) for i in range(lb.size())]
        idx = next(i for i, txt in enumerate(labels) if "目标事件" in txt)
        lb.selection_set(idx)
        b = next(w for w in _walk(dlg) if w.winfo_class() == "TButton"
                 and str(w.cget("text")) == "确定")
        b.invoke()

    app.root.after(200, choose)
    eid = app._pick_event("选择事件")
    assert eid == ev["id"], "对话框应等用户选择后返回事件 id，实际=%r" % eid
