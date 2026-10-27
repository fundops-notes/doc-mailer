"""Mail group placeholders and the pipeline's unified vertical scrolling."""
from pathlib import Path
import sys,tkinter as tk
from tkinter import ttk
import pytest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
import lp_v3,lp_tool


def test_mail_group_fields_resolve_actual_unit_and_single_fallback():
    rec={'name':'Investor A','group':'Group A','owner_pw':'secret-a','fields':{'合并组':'fake'}}
    ev={'recipients':[rec], 'options':{}}
    group=lp_v3.build_mail_field_map({'rep':rec,'group':'Group A'},ev)
    single=lp_v3.build_mail_field_map({'rep':rec,'group':''},ev)
    assert group['{{合并组}}']=='Group A'
    assert group['{{合并组或收件人名称}}']=='Group A'
    assert single['{{合并组}}']=='' and single['{{合并组或收件人名称}}']=='Investor A'
    assert 'fake' not in group['{{合并组}}']
    protected=lp_v3.build_mail_field_map({'rep':rec,'group':'secret-a'},ev)
    assert 'secret-a' not in protected['{{合并组}}']


def test_real_engine_saves_fake_drafts_with_group_subject_and_body(tmp_path,monkeypatch):
    import win32com.client
    class Attachments:
        def Add(self,path):return None
    class Mail:
        def __init__(self):self.Attachments=Attachments();self.To='';self.Subject='';self.HTMLBody=''
        def Save(self):saved.append(self)
    class Outlook:
        def CreateItem(self,_):return Mail()
    saved=[];monkeypatch.setattr(win32com.client,'Dispatch',lambda _:Outlook())
    src=tmp_path/'source.pdf';src.write_bytes(b'example')
    ev=lp_v3.new_event('Group test')
    ev['recipients']=[lp_v3.normalize_recipient({'name':name,'to':name+'@example.com',
        'group':group,'open_pw':'open-'+name,'owner_pw':'owner-'+name})
        for name,group in [('A','Group A'),('B','Group A'),('C','')]]
    ev['options']['send_mode']='per_group';item=lp_v3.new_item('pdf');item['params']['files']=[str(src)];ev['pipeline']=[item]
    ev['mail']['subject']='Report {{合并组或收件人名称}}'
    ev['mail']['body']='Group: {{合并组}}; recipient: {{合并组或收件人名称}}'
    store=lp_v3.Store(tmp_path/'store');store.save_settings({'outdir':str(tmp_path/'out')})
    result=lp_v3.run_event(store,ev,create_drafts=True,outdir=tmp_path/'out')
    assert result['status']=='ok' and len(saved)==2
    assert {m.Subject for m in saved}=={'Report Group A','Report C'}
    assert any('Group: Group A' in m.HTMLBody for m in saved)
    assert all('owner-' not in m.Subject+m.HTMLBody for m in saved)


@pytest.fixture
def app(tmp_path,monkeypatch,tk_root):
    monkeypatch.setattr(lp_v3,'default_base_dir',lambda:tmp_path)
    monkeypatch.setattr(lp_tool.App,'check_update',lambda *a,**k:None)
    monkeypatch.setattr(lp_tool.messagebox,'showinfo',lambda *a,**k:None)
    monkeypatch.setattr(lp_tool.messagebox,'showwarning',lambda *a,**k:None)
    a=lp_tool.App(tk_root)
    ev=lp_v3.new_event('Layout test')
    ev['recipients']=[lp_v3.normalize_recipient({'name':f'R{i}','to':f'r{i}@example.com',
        'open_pw':f'o{i}','owner_pw':f'owner{i}'}) for i in range(18)]
    it=lp_v3.new_item('excel');ev['pipeline']=[it];a.store.save_event(ev);a.open_event(ev['id'])
    tk_root.deiconify();tk_root.geometry('960x620');tk_root.update()
    yield a
    a._mail_dirty=False;a._ov_dirty=False;a.shutdown();tk_root.withdraw()


def walk(w):
    yield w
    for child in w.winfo_children():yield from walk(child)


def test_pipeline_has_one_vertical_scrollbar_and_all_table_rows(app):
    app._wb_show_step(1)
    for _ in range(8):app.root.update()
    vertical=[w for w in walk(app._wb_content) if isinstance(w,ttk.Scrollbar) and str(w.cget('orient'))=='vertical']
    assert vertical==[app._wb_pipeline_view.scrollbar]
    assert app._wb_params_view is app._wb_pipeline_view and app._wb_right_view is app._wb_pipeline_view
    assert int(app._wb_excel_table.cget('height'))>=18 and int(app._wb_map_tree.cget('height'))>=18
    assert app._wb_pipeline_view.canvas.yview()[1]<1
    app._wb_pipeline_view.canvas.yview_moveto(1);app.root.update()
    assert app._wb_pipeline_view.canvas.yview()[1]==pytest.approx(1)


def test_subject_and_body_field_menus_insert_group(app):
    app._wb_show_step(2);app.root.update()
    menus=[w for w in walk(app._wb_content) if isinstance(w,ttk.Menubutton)]
    subject=next(w for w in menus if w.cget('text')=='字段 ▾')
    menu=subject.nametowidget(subject.cget('menu'))
    i=next(i for i in range(menu.index('end')+1) if menu.entrycget(i,'label')=='{{合并组或收件人名称}}')
    menu.invoke(i)
    assert '{{合并组或收件人名称}}' in app.mail_subject_var.get() and app._mail_dirty
    body=next(w for w in menus if str(w.cget('text')).startswith('插入字段'))
    menu=body.nametowidget(body.cget('menu'))
    i=next(i for i in range(menu.index('end')+1) if menu.entrycget(i,'label')=='{{合并组}}')
    menu.invoke(i)
    assert '{{合并组}}' in app.body_text.get('1.0','end')


def test_excel_mapping_export_scroll_and_selection_position(app, monkeypatch):
    app._wb_show_step(1)
    for _ in range(5): app.root.update()
    item=app._wb_event['pipeline'][0]
    item['params']['path_source']='item.paths'
    item['params']['path_map']={'R0':['C:/reports/a.pdf','C:/reports/b.pdf']}
    app._wb_render_params(item)
    for _ in range(5): app.root.update()
    table=app._wb_excel_table
    assert table.xview()[1] < 1
    bar=next(w for w in table.master.winfo_children() if isinstance(w,ttk.Scrollbar))
    assert int(bar.grid_info()['row']) < int(table.grid_info()['row'])
    table.xview_moveto(1); app.root.update()
    assert table.xview()[0] > 0
    exported=[]
    monkeypatch.setattr(app,'_write_xlsx_sheets',lambda *args:exported.append(args))
    button=next(w for w in walk(app._wb_content) if isinstance(w,ttk.Button) and w.cget('text')=='导出当前对应表')
    button.invoke()
    rows=exported[0][1][0][1]
    assert rows[0]==['收件人名称','附件路径'] and rows[1]==['R0','C:/reports/a.pdf;C:/reports/b.pdf']
    view=app._wb_pipeline_view.canvas
    view.yview_moveto(0);app.root.update()
    app._wb_params_item_id='other';app._wb_render_params(item)
    for _ in range(6):app.root.update()
    assert view.yview()[0] < 0.1
