from pathlib import Path
import tkinter as tk
import time
from tkinter import ttk
import fitz
import pytest
from test_v3_8_4_layout import visible
import lp_tool,lp_v3
from test_v3_pool_gui import app,no_messagebox,_event,_walk


def test_item_watermark_config_roundtrip_and_legacy_fallback(tmp_path, monkeypatch):
    src=tmp_path/'test.pdf';doc=fitz.open();doc.new_page();doc.save(str(src));doc.close()
    event=lp_v3.new_event('legacy');rec=lp_v3.normalize_recipient({'name':'LP'});event['recipients']=[rec]
    event['options']['wm_settings']={'opacity':0.17}
    item=lp_v3.new_item('pdf');item['params'].update({'files':[str(src)],'watermark':True,'encrypt':False})
    event['pipeline']=[item];unit={'rep':rec,'recipients':[rec],'lp_name':'LP'};seen=[]
    monkeypatch.setattr(lp_v3.wm,'watermark_pdf',lambda s,o,t,c:(seen.append(dict(c or {})) or {'warnings':[]}))
    lp_v3.execute_item(item,unit,event,lp_v3.RunContext(event,tmp_path/'out'));assert seen[-1]['opacity']==0.17
    before=lp_v3.processing_signature(event);item['params']['wm_settings']={'opacity':0.34}
    normalized=lp_v3.normalize_event(event);assert normalized['pipeline'][0]['params']['wm_settings']=={'opacity':0.34}
    lp_v3.execute_item(normalized['pipeline'][0],unit,normalized,lp_v3.RunContext(normalized,tmp_path/'out'));assert seen[-1]['opacity']==0.34
    assert lp_v3.processing_signature(normalized)!=before
    other=lp_v3.new_item('pdf');other['params'].update({'files':[str(src)],'watermark':True});normalized['pipeline'].append(other)
    lp_v3.execute_item(other,unit,normalized,lp_v3.RunContext(normalized,tmp_path/'out'));assert seen[-1]['opacity']==0.17


def test_watermark_settings_follow_selected_pipeline_item(app):
    e=_event(app,'Watermark UI',[{'name':'LP','to':'lp@example.com'}]);it=lp_v3.new_item('pdf');e['pipeline']=[it];e['options']['wm_settings']={'opacity':0.20};app.store.save_event(e)
    app.show_screen('settings');app.root.update()
    assert not any(isinstance(w,ttk.Button) and w.cget('text')=='水印设置…' for w in _walk(app._screens['settings']))
    app.open_event(e['id']);app.root.deiconify();app.root.geometry('960x620');app._wb_show_step(1)
    for _ in range(5):app.root.update()
    checkbox=next(w for w in _walk(app._wb_content) if isinstance(w,ttk.Checkbutton) and w.cget('text')=='加水印')
    checkbox.invoke()
    for _ in range(4):app.root.update()
    button=next(w for w in _walk(app._wb_content) if isinstance(w,ttk.Button) and w.cget('text')=='水印设置…')
    button.invoke();app.root.update()
    assert app._wm_vars['opacity'].get()==0.20
    app._wm_vars['opacity'].set(0.35)
    win=next(w for w in app.root.winfo_children() if isinstance(w,tk.Toplevel) and w.title().startswith('水印设置'))
    save=next(w for w in _walk(win) if isinstance(w,ttk.Button) and w.cget('text')=='保存')
    assert save.winfo_ismapped() and save.winfo_rooty()+save.winfo_height() <= win.winfo_rooty()+win.winfo_height()
    save.invoke()
    for _ in range(4):app.root.update()
    saved,err=app.store.load_event(e['id']);assert not err
    assert saved['pipeline'][0]['params']['wm_settings']['opacity']==0.35
    assert saved['options']['wm_settings']['opacity']==0.20


def test_real_pdf_watermark_items_keep_independent_settings(tmp_path):
    src=tmp_path/'input.pdf';doc=fitz.open();page=doc.new_page();page.insert_text((72,72),'Evidence');doc.save(str(src));doc.close()
    e=lp_v3.new_event('Real watermark');r=lp_v3.normalize_recipient({'name':'Investor A','wm':'Investor A'});e['recipients']=[r]
    unit={'rep':r,'recipients':[r],'lp_name':'Investor A'};items=[]
    for opacity in (0.15,0.45):
        it=lp_v3.new_item('pdf');it['params'].update({'files':[str(src)],'watermark':True,'encrypt':False,'wm_settings':{'opacity':opacity}});items.append(it)
    e['pipeline']=items;outputs=[]
    for it in items:
        result,_=lp_v3.execute_item(it,unit,e,lp_v3.RunContext(e,tmp_path/'output'));outputs.append(result[0])
    assert outputs[0]!=outputs[1]
    with fitz.open(outputs[0]) as a, fitz.open(outputs[1]) as b:
        pix_a=a[0].get_pixmap().samples;pix_b=b[0].get_pixmap().samples
        assert pix_a!=pix_b and 'Evidence' in a[0].get_text() and 'Evidence' in b[0].get_text()


def test_checkbox_keyboard_focus_and_compact_buttons(app):
    e=_event(app,'Keyboard Watermark',[{'name':'LP','to':'lp@example.com'}]);it=lp_v3.new_item('pdf');e['pipeline']=[it];app.store.save_event(e);app.open_event(e['id'])
    app.root.deiconify();app.root.geometry('960x620');app._wb_show_step(1)
    for _ in range(5):app.root.update()
    check=app._wb_watermark_check;app._wb_pipeline_view.ensure_visible(check)
    for _ in range(4):app.root.update()
    check.focus_force();app.root.update();check.event_generate('<space>')
    for _ in range(6):app.root.update()
    assert app._wb_item_by_id(it['id'])['params']['watermark']
    new_check=app._wb_watermark_check;assert app.root.focus_get()==new_check
    new_check.event_generate('<Tab>')
    for _ in range(3):app.root.update()
    focused=app.root.focus_get();assert isinstance(focused,ttk.Button) and focused.cget('text')=='水印设置…'
    assert 22 <= focused.winfo_height() <= 38


@pytest.mark.parametrize('kind',['generate','pdf','file','excel'])
@pytest.mark.parametrize('size',['960x620','1280x800'])
def test_watermark_controls_reachable_across_sizes(app,kind,size):
    e=_event(app,'Watermark reachability',[{'name':'LP','to':'lp@example.com'}]);it=lp_v3.new_item(kind);it['params']['watermark']=True;e['pipeline']=[it];app.store.save_event(e);app.open_event(e['id'])
    app.root.deiconify();app.root.geometry(size);app._wb_show_step(1)
    for _ in range(5):app.root.update()
    button=next(w for w in _walk(app._wb_content) if isinstance(w,ttk.Button) and w.cget('text')=='水印设置…')
    deadline=time.monotonic()+3
    while time.monotonic()<deadline:
        app._wb_pipeline_view.ensure_visible(button, settle=0)
        app.root.update()
        if visible(button):
            break
    assert visible(button),(kind,size)
    for label in _walk(app._wb_params_frame):
        if isinstance(label,ttk.Label) and getattr(label,'_wrap_bound',False):
            assert label.winfo_width() >= label.winfo_reqwidth()-2, (kind,size,label.cget('text'))
    button.invoke();app.root.update()
    win=next(w for w in app.root.winfo_children() if isinstance(w,tk.Toplevel) and w.title().startswith('水印设置'))
    for label in ['保存','取消','刷新预览']:
        control=next(w for w in _walk(win) if isinstance(w,ttk.Button) and w.cget('text')==label)
        assert visible(control),(kind,size,label)
    win.destroy()


def test_watermark_dialog_accessible_at_larger_tk_scaling(app):
    old=app.root.tk.call('tk','scaling')
    try:
        app.root.tk.call('tk','scaling',1.5)
        e=_event(app,'Scaled Watermark',[{'name':'LP','to':'lp@example.com'}]);it=lp_v3.new_item('pdf');it['params']['watermark']=True;e['pipeline']=[it];app.store.save_event(e);app.open_event(e['id'])
        app.root.deiconify();app.root.geometry('960x620');app._wb_show_step(1)
        for _ in range(5):app.root.update()
        button=next(w for w in _walk(app._wb_content) if isinstance(w,ttk.Button) and w.cget('text')=='水印设置…')
        app._wb_pipeline_view.ensure_visible(button)
        for _ in range(4):app.root.update()
        assert visible(button)
        button.invoke();app.root.update()
        win=next(w for w in app.root.winfo_children() if isinstance(w,tk.Toplevel) and w.title().startswith('水印设置'))
        for label in ['保存','取消','刷新预览']:
            control=next(w for w in _walk(win) if isinstance(w,ttk.Button) and w.cget('text')==label)
            assert visible(control),label
        win.destroy()
    finally:app.root.tk.call('tk','scaling',old)
