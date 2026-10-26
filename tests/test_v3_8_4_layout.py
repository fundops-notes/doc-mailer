"""Real widget reachability at supported window sizes, including form rebuilds."""
import tkinter as tk
import pytest
import lp_tool
import lp_v3
from test_v3_pool_gui import app, no_messagebox, _walk, _event


def settle(a):
    for _ in range(5):
        a.root.update()


def visible(widget):
    if not widget.winfo_ismapped():
        return False
    x, y = widget.winfo_rootx(), widget.winfo_rooty()
    right, bottom = x + widget.winfo_width(), y + widget.winfo_height()
    parent = widget.master
    while parent is not None:
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        if x < px-2 or y < py-2 or right > px+parent.winfo_width()+2 or bottom > py+parent.winfo_height()+2:
            return False
        parent = parent.master
    if widget.winfo_class() in ('TButton','TCheckbutton','TMenubutton'):
        return widget.winfo_width() >= widget.winfo_reqwidth() - 2
    return True


def reveal(widget):
    parent = widget.master
    while parent is not None:
        view = getattr(parent, '_scroll_view', None)
        if view is not None:
            view.ensure_visible(widget)
        parent = parent.master


@pytest.fixture
def layout_app(app, monkeypatch):
    monkeypatch.setattr(app, 'check_update', lambda *a, **kw: None)
    callbacks = []
    monkeypatch.setattr(app.root, 'report_callback_exception', lambda *args: callbacks.append(args))
    app.root.deiconify()
    app.root.state('normal')
    ev = lp_v3.new_event('布局回归事件')
    ev['recipients'] = [lp_v3.normalize_recipient({'name':'测试收件人','to':'layout@example.com','open_pw':'test-open','owner_pw':'test-owner','fields':{'金额':'12'}})]
    attach = lp_v3.new_item('attach')
    attach['params'].update({'encrypt':True,'watermark':True,'files':['C:/测试/'+ '季度报告_'*40 + '.pdf']})
    generate = lp_v3.new_item('generate')
    generate['params'].update({'encrypt':True,'to_pdf':True})
    generate['mapping'] = {'{{字段%d}}'%i:'fields.金额' for i in range(24)}
    ev['pipeline'] = [attach, generate]
    ev['options'].update({'gp_owner_mode':True,'gp_owner_pw':'test-sender-owner'})
    app.store.save_event(ev)
    app.open_event(ev['id'])
    settle(app)
    yield app, attach['id'], generate['id']
    assert not callbacks, callbacks
    app._mail_dirty = app._ov_dirty = False
    app.root.state('normal')
    app.root.withdraw()


def choose(a, iid):
    a._wb_pipe_tree.selection_set(iid)
    a._wb_pipe_tree.event_generate('<<TreeviewSelect>>')
    settle(a)


@pytest.mark.parametrize('size', ['960x620','1280x800','1536x864'])
@pytest.mark.parametrize('kind', ['attach','generate'])
def test_all_parameter_controls_reachable(layout_app, size, kind):
    a, attach, generate = layout_app
    a.root.geometry(size)
    a._wb_show_step(1)
    choose(a, attach if kind == 'attach' else generate)
    view = a._wb_params_view
    reveal(view.scrollbar)
    settle(a)
    assert visible(view.scrollbar)
    assert visible(a._wb_prev_btn) and visible(a._wb_next_btn)
    for widget in _walk(a._wb_params_frame):
        if widget.winfo_class() in ('TButton','TCheckbutton','TEntry','TCombobox'):
            view.ensure_visible(widget)
            settle(a)
            assert visible(widget), (kind,size,widget)
    view.canvas.yview_moveto(1)
    settle(a)
    assert view.canvas.yview()[1] == pytest.approx(1)


def test_switching_item_rebuilds_scroll_region_and_resets_top(layout_app):
    a, attach, generate = layout_app
    a.root.geometry('960x620')
    a._wb_show_step(1)
    choose(a, generate)
    view = a._wb_params_view
    assert view.content.winfo_height() >= view.content.winfo_reqheight() - 1
    view.canvas.yview_moveto(1)
    choose(a, attach)
    assert view.canvas.yview()[0] == pytest.approx(0)
    choose(a, generate)
    assert view.content.winfo_height() >= view.content.winfo_reqheight() - 1
    assert view.content.winfo_height() > view.canvas.winfo_height()


def test_wheel_and_focus_reveal_controls_without_global_scroll_binding(layout_app):
    a, _, generate = layout_app
    a.root.geometry('960x620')
    a._wb_show_step(1)
    choose(a, generate)
    view = a._wb_params_view
    view.canvas.yview_moveto(0)
    view.canvas.event_generate('<MouseWheel>',delta=-120)
    settle(a)
    assert view.canvas.yview()[0] > 0
    target = next(w for w in _walk(a._wb_params_frame) if w.winfo_class()=='TButton' and w.cget('text')=='删除选中')
    target.event_generate('<FocusIn>')
    settle(a)
    assert visible(target)
    before = view.canvas.yview()
    a._nav_buttons['help'].event_generate('<MouseWheel>',delta=-120)
    settle(a)
    assert view.canvas.yview() == before


def test_scroll_bindings_are_removed_on_destroy(layout_app):
    a, _, _ = layout_app
    before = a.root.bind('<MouseWheel>')
    for _ in range(10):
        view = lp_tool.ScrollFrame(a._wb_content)
        callbacks=[binding for _sequence,binding in view._bindings]
        view.destroy()
        assert all(not a.root.tk.call('info','commands',binding) for binding in callbacks)
    active=lambda script:[line for line in script.splitlines() if line.strip()]
    assert active(a.root.bind('<MouseWheel>')) == active(before)


def test_path_text_wheel_does_not_scroll_both_text_and_form(layout_app):
    a, attach, _ = layout_app
    a.root.geometry('960x620')
    a._wb_show_step(1)
    choose(a, attach)
    text=next(w for w in _walk(a._wb_params_frame) if w.winfo_class()=='Text')
    text.insert('end','\n'.join('C:/test/file%d.pdf'%i for i in range(50)))
    a._wb_params_view.ensure_visible(text)
    settle(a)
    before=a._wb_params_view.canvas.yview()
    text.event_generate('<MouseWheel>',delta=-120)
    settle(a)
    assert text.yview()[0]>0
    assert a._wb_params_view.canvas.yview()==before


def test_path_box_has_working_horizontal_scroll_and_full_width_tree_bar(layout_app):
    a, attach, _ = layout_app
    a.root.geometry('960x620')
    a._wb_show_step(1)
    choose(a, attach)
    text = next(w for w in _walk(a._wb_params_frame) if w.winfo_class()=='Text')
    assert str(text.cget('xscrollcommand'))
    a._wb_params_view.ensure_visible(text)
    text.xview_moveto(1)
    settle(a)
    assert text.xview()[0] > 0
    assert a._wb_pipe_hsb.winfo_width() >= a._wb_pipe_tree.winfo_width() - 2


@pytest.mark.parametrize('page', ['events','pool','runs','settings','help',0,2,3])
def test_minimum_window_actions_are_reachable(layout_app, page):
    a, _, _ = layout_app
    a.root.geometry('960x620')
    if isinstance(page,int):
        a._wb_show_step(page)
        host = a._wb_content
    elif page in ('events','pool','runs'):
        a.show_home_section(page)
        host = {'events':a.home_tab_ev,'pool':a.home_tab_pool,'runs':a.home_tab_runs}[page]
    else:
        a.show_screen(page)
        host = a._screens[page]
    settle(a)
    for widget in list(_walk(host)):
        if widget.winfo_class() in ('TButton','TCheckbutton','TMenubutton','TCombobox'):
            reveal(widget)
            settle(a)
            assert visible(widget), (page,widget)


def test_permission_change_persists_after_scrolling_and_resizing(layout_app):
    a, attach, _ = layout_app
    a.root.geometry('960x620')
    a._wb_show_step(1)
    choose(a, attach)
    button=next(w for w in _walk(a._wb_params_frame) if w.winfo_class()=='TCheckbutton' and w.cget('text')=='允许复制内容')
    a._wb_params_view.ensure_visible(button)
    settle(a)
    assert visible(button)
    button.invoke()
    a.root.geometry('1280x800')
    settle(a)
    saved=a.store.load_event(a._wb_event['id'])[0]
    item=next(i for i in saved['pipeline'] if i['id']==attach)
    assert 'copy' in item['params']['permissions']
    assert saved['options']['gp_owner_pw']=='test-sender-owner'
    assert saved['recipients'][0]['owner_pw']=='test-owner'


def test_process_reveals_progress_log_on_current_page(layout_app, monkeypatch, tmp_path):
    a, attach, _ = layout_app
    a.root.geometry('960x620')
    a._wb_show_step(1)
    choose(a,attach)
    monkeypatch.setattr(lp_v3,'preflight',lambda *args:([],[],{}))
    monkeypatch.setattr(a,'_submit_job',lambda _job:None)
    monkeypatch.setattr(a,'_default_outdir',lambda:tmp_path/'output')
    a._wb_process_btn.invoke()
    settle(a)
    assert a._wb_step==1
    assert visible(a._wb_pipe_log)
    assert str(a._wb_pipe_log.cget('yscrollcommand'))
    a.running=False
