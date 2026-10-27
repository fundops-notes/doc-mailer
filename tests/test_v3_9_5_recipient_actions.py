"""Event recipient buttons must add, edit and copy independently."""
from pathlib import Path
import sys
import tkinter as tk
from tkinter import ttk

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import lp_v3
import lp_tool


@pytest.fixture
def app(tmp_path, monkeypatch, tk_root):
    monkeypatch.setattr(lp_v3, 'default_base_dir', lambda: tmp_path)
    notices = []
    for name in ('showinfo', 'showwarning', 'showerror'):
        monkeypatch.setattr(lp_tool.messagebox, name, lambda *a, **k: notices.append(a))
    monkeypatch.setattr(lp_tool.App, 'check_update', lambda *a, **k: None)
    a = lp_tool.App(tk_root)
    event = lp_v3.new_event('recipient action test')
    event['recipients'] = [lp_v3.normalize_recipient({'name':'Alice','to':'alice@example.com',
        'wm':'Alice watermark','open_pw':'open-alice','owner_pw':'owner-alice',
        'fields':{'source':'keep'}})]
    assert a.store.save_event(event)[0]
    a.open_event(event['id']); a._wb_show_step(0); tk_root.update()
    yield a, notices, event['id']
    for win in tk_root.winfo_children():
        if isinstance(win, tk.Toplevel): win.destroy()
    a.shutdown()


def walk(w):
    yield w
    for child in w.winfo_children(): yield from walk(child)


def btn(w, label):
    return next(c for c in walk(w) if isinstance(c, ttk.Button) and c.cget('text') == label)


def dialog(a, title):
    return next(w for w in a.root.winfo_children() if isinstance(w, tk.Toplevel) and w.title() == title)


def entries(win):
    return [w for w in win.winfo_children() if isinstance(w, ttk.Entry)]


def select_first(a):
    tree = a._wb_rec_tree; tree.selection_set(tree.get_children()[0]); a.root.update()


def saved(a, event_id):
    return a.store.load_event(event_id)[0]['recipients']


def test_add_ignores_selection_and_keeps_existing_recipient(app):
    a, notices, event_id = app; select_first(a)
    before = saved(a, event_id)[0]
    btn(a._wb_content, '新增').invoke(); win = dialog(a, '新增收件人')
    fields = entries(win); assert fields[0].get() == '' and fields[2].get() == ''
    fields[0].insert(0, 'Bob'); fields[2].insert(0, 'bob@example.com')
    btn(win, '保存').invoke(); a.root.update()
    result = saved(a, event_id)
    assert len(result) == 2 and result[0] == before and result[1]['name'] == 'Bob'


def test_copy_creates_independent_recipient_preserving_fields(app):
    a, notices, event_id = app; select_first(a)
    before = saved(a, event_id)[0]
    btn(a._wb_content, '复制').invoke(); win = dialog(a, '复制收件人')
    fields = entries(win)
    assert fields[0].get() == 'Alice（副本）' and fields[2].get() == 'alice@example.com'
    assert any('核对' in str(w.cget('text')) for w in win.winfo_children() if isinstance(w, ttk.Label))
    fields[2].delete(0, 'end'); fields[2].insert(0, 'copy@example.com')
    btn(win, '保存').invoke(); a.root.update()
    result = saved(a, event_id)
    assert len(result) == 2 and result[0] == before
    assert result[1]['name'] == 'Alice（副本）' and result[1]['to'] == 'copy@example.com'
    assert result[1]['fields'] == before['fields'] and result[1]['owner_pw'] == before['owner_pw']


def test_copy_cancel_and_duplicate_name_are_safe(app):
    a, notices, event_id = app; select_first(a)
    before = saved(a, event_id)
    btn(a._wb_content, '复制').invoke(); win = dialog(a, '复制收件人')
    name = entries(win)[0]; name.delete(0, 'end'); name.insert(0, 'Alice')
    btn(win, '保存').invoke()
    assert win.winfo_exists() and saved(a, event_id) == before and any(x[0] == '名称重复' for x in notices)
    btn(win, '取消').invoke()
    assert saved(a, event_id) == before


def test_edit_and_copy_require_exactly_one_selection(app):
    a, notices, event_id = app
    for text in ('编辑', '复制'):
        btn(a._wb_content, text).invoke()
    assert sum(n[0] == '请选择收件人' for n in notices) == 2 and saved(a, event_id)[0]['name'] == 'Alice'


def test_edit_only_changes_selected_recipient(app):
    a, notices, event_id = app; select_first(a)
    btn(a._wb_content, '编辑').invoke(); win = dialog(a, '编辑收件人')
    to = entries(win)[2]; to.delete(0, 'end'); to.insert(0, 'updated@example.com')
    btn(win, '保存').invoke(); a.root.update()
    result = saved(a, event_id)
    assert len(result) == 1 and result[0]['to'] == 'updated@example.com' and result[0]['fields'] == {'source':'keep'}


def test_failed_save_does_not_leave_unsaved_duplicate(app, monkeypatch):
    a, notices, event_id = app; select_first(a)
    before = saved(a, event_id)
    btn(a._wb_content, '复制').invoke(); win = dialog(a, '复制收件人')
    monkeypatch.setattr(a, '_wb_save', lambda: False)
    btn(win, '保存').invoke()
    assert win.winfo_exists() and a._wb_event['recipients'] == before and saved(a, event_id) == before
    btn(win, '取消').invoke()
