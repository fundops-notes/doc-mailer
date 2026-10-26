"""Toolbar field insertion must participate in unsaved-edit guards."""
import pytest
import lp_tool
from test_v3_pool_gui import app, no_messagebox
from test_v3_9_0_workflow import make_event, show, texts


@pytest.mark.parametrize('step', [2, 3])
def test_field_menu_marks_active_editor_dirty(app, tmp_path, step):
    event, _ = make_event(tmp_path)
    show(app, event, step)
    menu_button = texts(app._editor_toolbar)['插入字段 ▾']
    menu = app.root.nametowidget(str(menu_button['menu']))
    menu.invoke(0)
    assert app._mail_dirty if step == 2 else app._ov_dirty


def test_field_insertion_requires_save_confirmation_before_preview(app, tmp_path, monkeypatch):
    event, _ = make_event(tmp_path)
    show(app, event, 2)
    saved = app.store.event_path(event['id']).read_bytes()
    app._insert_editor_text('{{收件人名称}}')
    prompts = []
    monkeypatch.setattr(lp_tool.messagebox, 'askokcancel', lambda *args, **kw: prompts.append(args) or False)
    app._wb_next_btn.invoke()
    assert prompts and app._wb_step == 2 and app._mail_dirty
    assert app.store.event_path(event['id']).read_bytes() == saved
