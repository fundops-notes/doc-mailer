import hashlib
import lp_tool, lp_v3
from test_v3_pool_gui import app, no_messagebox, _event


def test_empty_pipeline_single_and_batch_save_only(tmp_path, monkeypatch):
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    ev=lp_v3.new_event('No files')
    ev['recipients']=[lp_v3.normalize_recipient({'name':'Alice','to':'alice@example.com'})]
    ev['pipeline']=[]
    ev['mail']['subject']='Hello {{收件人名称}}'
    errors,warnings,_=lp_v3.preflight(ev,store)
    assert not errors and any('流水线为空' in x for x in warnings)
    assert len(lp_v3.preview_plan(ev)) == 1
    saved=[]
    class Attachments:
        def Add(self,p): raise AssertionError('unexpected attachment')
    class Accessor:
        def __init__(self): self.values={}
        def SetProperty(self,key,value): self.values[key]=value
        def GetProperty(self,key): return self.values[key]
    class Mail:
        def __init__(self):self.Attachments=Attachments();self.PropertyAccessor=Accessor();self.Subject='';self.HTMLBody=''
        def Save(self):saved.append((self.Subject,self.HTMLBody))
    class Outlook:
        def CreateItem(self,n):return Mail()
    monkeypatch.setattr('win32com.client.Dispatch',lambda _:Outlook())
    run=lp_v3.run_event(store,ev,outdir=tmp_path/'out',word=object())
    assert run['status']=='ok' and run['items'][0]['files']==[] and len(saved)==1
    assert 'Alice' in saved[0][0]
    assert store.save_event(ev)[0]
    loaded,err=store.load_event(ev['id']);assert not err
    digest=hashlib.sha256(store.event_path(ev['id']).read_bytes()).hexdigest()
    result=lp_v3.run_batch_events(store,[(loaded,digest)],tmp_path/'out',object())
    assert result[0]['status']=='完成' and len(saved)==2


def test_no_pipeline_confirmation_can_cancel(app,monkeypatch):
    ev=_event(app,'Empty pipeline',[{'name':'Alice','to':'alice@example.com'}]);ev['pipeline']=[]
    assert app.store.save_event(ev)[0];app.open_event(ev['id']);app._wb_show_step(1)
    prompts=[]
    monkeypatch.setattr(lp_tool.messagebox,'askyesno',lambda title,body: prompts.append((title,body)) or False)
    assert app._wb_process() is False
    assert prompts and '流水线为空' in prompts[0][0] and not app.running


def test_empty_pipeline_next_reaches_mail_template(app,monkeypatch):
    ev=_event(app,'No attachment',[{'name':'Alice','to':'alice@example.com'}]);ev['pipeline']=[]
    assert app.store.save_event(ev)[0];app.open_event(ev['id']);app._wb_show_step(1)
    monkeypatch.setattr(lp_tool.messagebox,'askyesno',lambda *args,**kwargs:True)
    monkeypatch.setattr(app,'_submit_job',lambda job:job(lambda:object()))
    app._wb_next_btn.invoke()
    app._poll_queue()
    assert app._wb_step==2 and app._wb_last_run['status']=='ok'


def test_configured_pipeline_that_produces_nothing_still_skips(tmp_path,monkeypatch):
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    ev=lp_v3.new_event('No output')
    ev['recipients']=[lp_v3.normalize_recipient({'name':'Alice','to':'alice@example.com','open_pw':'abc','owner_pw':'xyz'})]
    it=lp_v3.new_item('pdf');it['params'].update(files=[],watermark=False,encrypt=False);ev['pipeline']=[it]
    monkeypatch.setattr(lp_v3,'preflight',lambda *args:([],[],{}))
    monkeypatch.setattr(lp_v3,'run_items',lambda *args,**kwargs:({it['id']:{}},[],[]))
    monkeypatch.setattr(lp_v3,'create_outlook_draft',lambda *args,**kwargs: (_ for _ in ()).throw(AssertionError('must not save')))
    run=lp_v3.run_event(store,ev,outdir=tmp_path/'out',word=object())
    assert run['items'][0]['draft']=='跳过(无附件)'
