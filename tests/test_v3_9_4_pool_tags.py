"""Recipient-library tag filtering and protected bulk replacement."""
import json
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
    monkeypatch.setattr(lp_tool.messagebox, 'showinfo', lambda *a, **k: None)
    monkeypatch.setattr(lp_tool.messagebox, 'showwarning', lambda *a, **k: None)
    monkeypatch.setattr(lp_tool.messagebox, 'showerror', lambda *a, **k: None)
    monkeypatch.setattr(lp_tool.messagebox, 'askyesno', lambda *a, **k: True)
    a = lp_tool.App(tk_root)
    yield a
    a.shutdown()


def make(store, name, tags):
    rec = lp_v3.new_pool_recipient({'name': name, 'to': name+'@example.com',
                                    'open_pw': 'open-'+name, 'owner_pw': 'owner-'+name,
                                    'tags': tags})
    ok, error, result = store.save_pool_recipient(rec)
    assert ok, error
    return result


def walk(w):
    yield w
    for child in w.winfo_children():
        yield from walk(child)


def button(win, label):
    return next(w for w in walk(win) if isinstance(w, ttk.Button) and w.cget('text') == label)


def popup(app):
    return [w for w in app.root.winfo_children() if isinstance(w, tk.Toplevel)][-1]


def visible_names(app):
    return {app._pool_tree.item(i, 'values')[0] for i in app._pool_tree.get_children()}


def toggle_menu(app, label):
    menu = app._pool_filter_menu
    for i in range((menu.index('end') or 0)+1):
        if menu.entrycget(i, 'label') == label:
            menu.invoke(i)
            return
    raise AssertionError('missing tag in dropdown: '+label)


def test_filter_dropdown_any_all_clear_and_selection(app):
    make(app.store, 'A', ['red', 'east'])
    make(app.store, 'B', ['red'])
    make(app.store, 'C', ['east'])
    make(app.store, 'D', [])
    app.show_home_section('pool'); app._refresh_pool_tab()
    assert visible_names(app) == {'A', 'B', 'C', 'D'}
    toggle_menu(app, 'red'); assert visible_names(app) == {'A', 'B'}
    toggle_menu(app, 'east'); assert visible_names(app) == {'A', 'B', 'C'}
    app._pool_filter_mode.set('匹配全部'); app._refresh_pool_tab()
    assert visible_names(app) == {'A'}
    app._pool_tree.selection_set(app._pool_tree.get_children())
    assert len(app._pool_selected()) == 1
    app._pool_clear_filter()
    assert visible_names(app) == {'A', 'B', 'C', 'D'}
    assert len(app._pool_selected()) == 1


def test_many_tags_dialog_scrolls_and_keeps_actions_visible(app):
    make(app.store, 'A', ['标签%03d' % i for i in range(100)])
    app.root.deiconify(); app.show_home_section('pool'); app._refresh_pool_tab()
    app._pool_tree.selection_set(app._pool_tree.get_children())
    app._pool_tags(); app.root.update(); win = popup(app)
    view = next(w for w in walk(win) if isinstance(w, lp_tool.ScrollFrame))
    assert view.canvas.yview()[1] < 1.0
    assert button(win, '替换勾选标签').winfo_rooty() + button(win, '替换勾选标签').winfo_height() <= win.winfo_rooty() + win.winfo_height()
    view.canvas.yview_moveto(1); app.root.update()
    last = next(w for w in walk(win) if isinstance(w, ttk.Checkbutton) and w.cget('text').startswith('标签099'))
    assert view.canvas.winfo_rooty() <= last.winfo_rooty() < view.canvas.winfo_rooty() + view.canvas.winfo_height()
    button(win, '关闭').invoke(); app.root.withdraw()


def test_batch_replace_through_real_controls_preserves_unselected(app):
    a = make(app.store, 'A', ['old', 'keep'])
    b = make(app.store, 'B', ['old', 'new'])
    c = make(app.store, 'C', ['old'])
    app.show_home_section('pool'); app._refresh_pool_tab()
    rows = [iid for iid in app._pool_tree.get_children()
            if app._pool_tree.item(iid, 'values')[0] in ('A', 'B')]
    app._pool_tree.selection_set(rows); app._pool_tags(); win = popup(app)
    next(w for w in walk(win) if isinstance(w, ttk.Checkbutton) and w.cget('text').startswith('old')).invoke()
    next(w for w in walk(win) if isinstance(w, ttk.Entry)).insert(0, 'new')
    button(win, '替换勾选标签').invoke(); app.root.update()
    records = {r['name']: r for r in app.store.list_pool()}
    assert records['A']['tags'] == ['new', 'keep']
    assert records['B']['tags'] == ['new']
    assert records['C']['tags'] == ['old']
    assert records['A']['owner_pw'] == a['owner_pw'] and records['B']['to'] == b['to']
    backups = list((app.store.data_dir/'backups').glob('pool_tags_*'))
    assert len(backups) == 1 and {p.name for p in backups[0].iterdir()} == {a['id']+'.json', b['id']+'.json'}


def test_global_rename_ignores_filter_and_merges_existing_tag(app):
    a = make(app.store, 'A', ['old', 'new', 'keep'])
    b = make(app.store, 'B', ['old'])
    make(app.store, 'C', ['other'])
    app.show_home_section('pool'); app._refresh_pool_tab()
    toggle_menu(app, 'old'); assert visible_names(app) == {'A', 'B'}
    app._pool_rename_tag(); win = popup(app)
    combo = next(w for w in walk(win) if isinstance(w, ttk.Combobox))
    combo.set('old'); combo.event_generate('<<ComboboxSelected>>'); app.root.update()
    next(w for w in walk(win) if isinstance(w, ttk.Entry) and not isinstance(w, ttk.Combobox)).insert(0, 'new')
    button(win, '重命名全部').invoke(); app.root.update()
    records = {r['name']: r for r in app.store.list_pool()}
    assert records['A']['tags'] == ['new', 'keep'] and records['B']['tags'] == ['new']
    assert records['C']['tags'] == ['other'] and visible_names(app) == {'A', 'B'}
    assert app._pool_filter_vars['new'].get()
    backup = next((app.store.data_dir/'backups').glob('pool_tags_*'))
    assert json.loads((backup/(a['id']+'.json')).read_text(encoding='utf-8'))['tags'] == ['old', 'new', 'keep']
    assert json.loads((backup/(b['id']+'.json')).read_text(encoding='utf-8'))['tags'] == ['old']


def test_global_rename_also_updates_records_hidden_by_filter(app):
    make(app.store, 'A', ['old', 'visible'])
    make(app.store, 'B', ['old'])
    app.show_home_section('pool'); app._refresh_pool_tab()
    toggle_menu(app, 'visible'); assert visible_names(app) == {'A'}
    app._pool_rename_tag(); win = popup(app)
    combo = next(w for w in walk(win) if isinstance(w, ttk.Combobox))
    combo.set('old')
    next(w for w in walk(win) if isinstance(w, ttk.Entry) and not isinstance(w, ttk.Combobox)).insert(0, 'new')
    button(win, '重命名全部').invoke(); app.root.update()
    records = {r['name']: r for r in app.store.list_pool()}
    assert records['A']['tags'] == ['new', 'visible'] and records['B']['tags'] == ['new']
    assert visible_names(app) == {'A'}


def test_replacement_preserves_unknown_fields_and_existing_bak(tmp_path):
    store = lp_v3.Store(tmp_path); a = make(store, 'A', ['old', 'keep', 'other'])
    path = store.pool_path(a['id']); raw = json.loads(path.read_text(encoding='utf-8'))
    raw['custom_future_field'] = {'value': 42}; path.write_text(json.dumps(raw), encoding='utf-8')
    before = path.read_bytes(); bak = path.with_name(path.name+'.bak'); bak.write_bytes(b'prior-backup')
    ok, error, stats = store.replace_pool_tags(['old', 'other'], 'new', [a['id']])
    assert ok, error
    after = json.loads(path.read_text(encoding='utf-8'))
    assert after['tags'] == ['new', 'keep'] and after['custom_future_field'] == {'value': 42}
    assert bak.read_bytes() == b'prior-backup'
    assert (Path(stats['backup'])/path.name).read_bytes() == before


def test_replacement_matches_trimmed_legacy_tags_without_losing_other_values(tmp_path):
    store = lp_v3.Store(tmp_path); rec = make(store, 'A', ['old', 'keep'])
    path = store.pool_path(rec['id']); data = json.loads(path.read_text(encoding='utf-8'))
    data['tags'] = [' old ', ' untouched ', 'new ']
    path.write_text(json.dumps(data), encoding='utf-8')
    ok, error, _ = store.replace_pool_tags(['old'], 'new')
    assert ok, error
    assert json.loads(path.read_text(encoding='utf-8'))['tags'] == [' untouched ', 'new ']


def test_backup_failure_does_not_modify_recipients(tmp_path, monkeypatch):
    store = lp_v3.Store(tmp_path); rec = make(store, 'A', ['old']); path = store.pool_path(rec['id'])
    before = path.read_bytes(); real_open = Path.open
    def deny_backup(self, *args, **kwargs):
        if self.parent.name.startswith('pool_tags_'):
            raise OSError('simulated backup error')
        return real_open(self, *args, **kwargs)
    monkeypatch.setattr(Path, 'open', deny_backup)
    ok, error, stats = store.replace_pool_tags(['old'], 'new')
    assert not ok and 'backup error' in error and path.read_bytes() == before and stats['changed'] == 0


def test_partial_write_failure_rolls_back_and_keeps_backups(tmp_path, monkeypatch):
    store = lp_v3.Store(tmp_path); a = make(store, 'A', ['old']); b = make(store, 'B', ['old'])
    paths = [store.pool_path(a['id']), store.pool_path(b['id'])]
    before = {p.name: p.read_bytes() for p in paths}; real = lp_v3.os.replace; calls = []
    def failing(src, dst):
        calls.append(str(dst))
        if len(calls) == 2:
            raise OSError('simulated second write failure')
        return real(src, dst)
    monkeypatch.setattr(lp_v3.os, 'replace', failing)
    ok, error, stats = store.replace_pool_tags(['old'], 'new', [a['id'], b['id']])
    assert not ok and 'second write failure' in error and not stats['rollback_incomplete']
    assert {p.name: p.read_bytes() for p in paths} == before
    assert all((Path(stats['backup'])/p.name).read_bytes() == before[p.name] for p in paths)


def test_temp_cleanup_failure_does_not_prevent_rollback(tmp_path, monkeypatch):
    store = lp_v3.Store(tmp_path); a = make(store, 'A', ['old']); b = make(store, 'B', ['old'])
    paths = [store.pool_path(a['id']), store.pool_path(b['id'])]
    before = [p.read_bytes() for p in paths]; real_replace = lp_v3.os.replace; real_unlink = Path.unlink
    calls = []
    def failing_replace(src, dst):
        calls.append(str(dst))
        if len(calls) == 2: raise OSError('simulated write failure')
        return real_replace(src, dst)
    def failing_unlink(path, *args, **kwargs):
        if path.name.endswith('.tags.tmp'): raise OSError('simulated cleanup failure')
        return real_unlink(path, *args, **kwargs)
    monkeypatch.setattr(lp_v3.os, 'replace', failing_replace)
    monkeypatch.setattr(Path, 'unlink', failing_unlink)
    ok, error, stats = store.replace_pool_tags(['old'], 'new', [a['id'], b['id']])
    assert not ok and 'write failure' in error and not stats['rollback_incomplete']
    assert [p.read_bytes() for p in paths] == before


def test_external_change_is_not_overwritten_and_first_write_restored(tmp_path, monkeypatch):
    store = lp_v3.Store(tmp_path); a = make(store, 'A', ['old']); b = make(store, 'B', ['old'])
    first, second = store.pool_path(a['id']), store.pool_path(b['id'])
    before = first.read_bytes(); changed = json.loads(second.read_text(encoding='utf-8'))
    changed['other_session'] = True; outside = json.dumps(changed).encode()
    real = lp_v3.os.replace; calls = []
    def concurrent(src, dst):
        real(src, dst); calls.append(str(dst))
        if len(calls) == 1: second.write_bytes(outside)
    monkeypatch.setattr(lp_v3.os, 'replace', concurrent)
    ok, error, stats = store.replace_pool_tags(['old'], 'new', [a['id'], b['id']])
    assert not ok and '其他操作修改' in error and first.read_bytes() == before
    assert second.read_bytes() == outside and not stats['rollback_incomplete']


def test_malformed_recipient_refuses_global_edit(tmp_path):
    store = lp_v3.Store(tmp_path); rec = make(store, 'A', ['old']); path = store.pool_path(rec['id'])
    before = path.read_bytes(); bad = store.recipients_dir/'bad.json'; bad.write_text('{bad', encoding='utf-8')
    ok, error, stats = store.replace_pool_tags(['old'], 'new')
    assert not ok and path.read_bytes() == before and bad.read_text(encoding='utf-8') == '{bad'
