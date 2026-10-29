import hashlib
from pathlib import Path
import fitz
import lp_tool,lp_v3
from test_v3_pool_gui import app,no_messagebox,_event,_walk
from tkinter import ttk

def test_font_sizes_save_points_and_preserve_legacy_pixels(tk_root):
    import tkinter as tk
    t=tk.Text(tk_root)
    try:
        lp_tool.load_html(t,'<div style="font-size:18px;font-family:Aptos">Legacy</div>')
        assert lp_tool._style_at(t,'1.0')['size_unit']=='px'
        assert 'font-size:18px' in lp_tool.text_to_html(t)
        t.tag_add('sel','1.0','1.3')
        class Editor:
            body_text=t
            _ed_need_sel=lambda self:('1.0','1.3')
            _mark_active_dirty=lambda self:None
        lp_tool.App._ed_restyle(Editor(),s=18)
        out=lp_tool.text_to_html(t)
        assert 'font-size:18pt' in out and 'font-size:18px' in out
        lp_tool.App._ed_restyle(Editor(),s=20)
        out=lp_tool.text_to_html(t)
        assert 'font-size:20pt' in out
        _,runs=lp_tool.html_to_runs(out)
        assert any(st.get('s')==20 and st.get('size_unit')=='pt' for _,_,st in runs)
    finally:t.destroy()

def test_batch_real_engine_fake_outlook_stops_incomplete_event(tmp_path,monkeypatch):
    src=tmp_path/'input.pdf'
    d=fitz.open();d.new_page();d.save(src);d.close()
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    events=[]
    for name in ('One','Two','Bad'):
        ev=lp_v3.new_event(name)
        ev['recipients']=[lp_v3.normalize_recipient({'name':name,'to':name+'@example.com','open_pw':'open-'+name,'owner_pw':'secret-'+name})]
        it=lp_v3.new_item('pdf');it['params'].update(files=[str(src if name!='Bad' else tmp_path/'missing.pdf')],watermark=False,encrypt=False)
        ev['pipeline']=[it]
        ev['mail']['subject']='Template {{收件人名称}}'
        ev['mail']['body']='<div style="font-size:18pt">Example <span style="font-size:20pt">body</span></div>'
        if name=='One':ev['mail_overrides']={'§One':{'subject':'Custom subject','body':'<p>Custom</p>'}}
        ok,err=store.save_event(ev);assert ok,err
        loaded,err=store.load_event(ev['id']);assert loaded,err
        events.append((loaded,hashlib.sha256(store.event_path(ev['id']).read_bytes()).hexdigest()))
    saved=[]
    class Attachments:
        def Add(self,p):assert Path(p).is_file()
    class Mail:
        def __init__(self):self.Attachments=Attachments();self.Subject='';self.HTMLBody=''
        def Save(self):saved.append((self.Subject,self.HTMLBody))
    class Outlook:
        def CreateItem(self,n):assert n==0;return Mail()
    monkeypatch.setattr('win32com.client.Dispatch',lambda _:Outlook())
    results=lp_v3.run_batch_events(store,events,tmp_path/'out',object())
    assert [r['status'] for r in results]==['完成','完成','跳过']
    assert len(saved)==2 and saved[0][0]=='Custom subject'
    assert 'font-size:18pt' in saved[1][1] and 'font-size:20pt' in saved[1][1]
    assert all('secret-' not in a+b for a,b in saved)
    for e,h in events[:2]:
        current,err=store.load_event(e['id']);assert current and not err
        assert current['mail_overrides']==e['mail_overrides']
        assert current['process_results'] and lp_v3.prepared_files_current(current)
        assert current['recipients']==e['recipients'] and current['pipeline']==e['pipeline']
        assert list(store.data_dir.joinpath('backups').glob(e['id']+'.before_batch_*.json'))
    assert hashlib.sha256(store.event_path(events[2][0]['id']).read_bytes()).hexdigest()==events[2][1]

def test_batch_selection_and_confirmation_uses_real_controls(app,monkeypatch,tmp_path):
    events=[]
    src=tmp_path/'file.pdf';d=fitz.open();d.new_page();d.save(src);d.close()
    for name in ('First','Second'):
        e=_event(app,name,[{'name':name,'to':name+'@example.com','open_pw':'open-'+name,'owner_pw':'owner-'+name}])
        it=lp_v3.new_item('pdf');it['params'].update(files=[str(src)],watermark=False,encrypt=False);e['pipeline']=[it];app.store.save_event(e);events.append(e)
    app._refresh_events_tab();app.ev_tree.selection_set([e['id'] for e in events]);app.root.update()
    assert len(app.ev_tree.selection())==2
    called=[]
    monkeypatch.setattr(app,'_submit_job',lambda job:called.append(job))
    monkeypatch.setattr(lp_tool.messagebox,'askyesno',lambda *args,**kwargs:True)
    app._batch_button.invoke()
    assert len(called)==1 and app.running
    app.running=False;app._batch_button.state(['!disabled']);app.ev_tree.state(['!disabled'])

def test_batch_skips_if_event_changes_during_processing(tmp_path,monkeypatch):
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    e=lp_v3.new_event('Concurrent');e['recipients']=[lp_v3.normalize_recipient({'name':'LP'})]
    ok,err=store.save_event(e);assert ok,err
    e,err=store.load_event(e['id']);path=store.event_path(e['id'])
    digest=hashlib.sha256(path.read_bytes()).hexdigest();calls=[]
    monkeypatch.setattr(lp_v3,'preflight',lambda *a:([],[],{}))
    def processing(*args,**kwargs):
        calls.append(kwargs['create_drafts'])
        path.write_text(path.read_text(encoding='utf-8')+' ',encoding='utf-8')
        return {'errors':[],'items':[],'stats':{'fail':0}}
    monkeypatch.setattr(lp_v3,'run_event',processing)
    result=lp_v3.run_batch_events(store,[(e,digest)],tmp_path,object())
    assert calls==[False] and result[0]['status']=='跳过'
    assert '处理期间发生变化' in result[0]['detail']

def test_real_size_control_roundtrips_into_fake_outlook(app,monkeypatch):
    e=_event(app,'Size UI',[{'name':'LP','to':'lp@example.com','open_pw':'open','owner_pw':'secret'}])
    app.open_event(e['id']);app.root.deiconify();app._wb_show_step(2)
    for _ in range(3):app.root.update()
    t=app.body_text;t.delete('1.0','end');t.insert('1.0','Size eighteen and twenty')
    t.tag_add('sel','1.0','1.4');app._ed_size.set('18');app._ed_size.event_generate('<<ComboboxSelected>>');app.root.update()
    t.tag_remove('sel','1.0','end');t.tag_add('sel','1.5','1.13')
    app._ed_size.set('20');app._ed_size.event_generate('<<ComboboxSelected>>');app.root.update()
    body=lp_tool.text_to_html(t)
    assert 'font-size:18pt' in body and 'font-size:20pt' in body
    class Mail:
        def __init__(self):self.Subject='';self.HTMLBody='';self.saved=False
        def Save(self):self.saved=True
    class Outlook:
        def CreateItem(self,n):return Mail()
    monkeypatch.setattr('win32com.client.Dispatch',lambda _:Outlook())
    r=lp_v3.normalize_recipient({'name':'LP','to':'lp@example.com','owner_pw':'secret'})
    mail=lp_v3.create_outlook_draft(r,[],'Subject',body,{},'secret')
    assert mail.saved and 'font-size:18pt' in mail.HTMLBody and 'font-size:20pt' in mail.HTMLBody

def test_batch_does_not_draft_if_file_disappears_during_processing(tmp_path,monkeypatch):
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    e=lp_v3.new_event('Missing during processing')
    e['recipients']=[lp_v3.normalize_recipient({'name':'LP'})]
    ok,err=store.save_event(e);assert ok,err
    e,err=store.load_event(e['id']);path=store.event_path(e['id'])
    digest=hashlib.sha256(path.read_bytes()).hexdigest();calls=[]
    monkeypatch.setattr(lp_v3,'preflight',lambda *a:([],[],{}))
    def processing(*args,**kwargs):
        calls.append(kwargs['create_drafts'])
        return {'errors':[],'warnings':['文件不存在，已跳过：example.pdf'],'items':[],'stats':{'fail':0}}
    monkeypatch.setattr(lp_v3,'run_event',processing)
    result=lp_v3.run_batch_events(store,[(e,digest)],tmp_path,object())
    assert calls==[False] and result[0]['status']=='处理失败'

def test_batch_does_not_create_drafts_when_event_backup_fails(tmp_path,monkeypatch):
    src=tmp_path/'input.pdf';d=fitz.open();d.new_page();d.save(src);d.close()
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    e=lp_v3.new_event('Backup fails')
    e['recipients']=[lp_v3.normalize_recipient({'name':'LP','to':'lp@example.com','open_pw':'open','owner_pw':'owner'})]
    item=lp_v3.new_item('pdf');item['params'].update(files=[str(src)],watermark=False,encrypt=False);e['pipeline']=[item]
    ok,err=store.save_event(e);assert ok,err
    e,err=store.load_event(e['id']);path=store.event_path(e['id']);before=path.read_bytes()
    def fail_backup(*_):raise OSError('simulated backup failure')
    monkeypatch.setattr(store,'_backup_event',fail_backup)
    calls=[]
    monkeypatch.setattr(lp_v3,'create_outlook_draft',lambda *a,**k:calls.append('unexpected'))
    result=lp_v3.run_batch_events(store,[(e,hashlib.sha256(before).hexdigest())],tmp_path/'out',object())
    assert result[0]['status']=='跳过' and '备份事件失败' in result[0]['detail']
    assert calls==[] and path.read_bytes()==before
