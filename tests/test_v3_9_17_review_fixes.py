"""Verify saved drafts and encoding warnings through actual event workflows."""
import hashlib
from pathlib import Path
import sys
import fitz
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import lp_v3


class Accessor:
    def __init__(self, mode): self.values={}; self.mode=mode
    def SetProperty(self,key,value):
        if self.mode=='set_fail': raise RuntimeError('set failed')
        self.values[key]=value
    def GetProperty(self,key):
        if self.mode=='read_fail': raise RuntimeError('read failed')
        if self.mode=='mismatch' and '3FFD' in key: return 936
        return self.values[key]


class Mail:
    def __init__(self,mode):
        self.mode=mode; self.pa=Accessor(mode); self.saved=0
        self.Attachments=type('Attachments',(),{'Add':lambda self,path:None})()
    @property
    def PropertyAccessor(self):
        if self.mode=='accessor_after_save' and self.saved: raise RuntimeError('accessor failure')
        return self.pa
    def Save(self): self.saved+=1


def event_fixture(tmp_path):
    source=tmp_path/'src.pdf'; doc=fitz.open();doc.new_page();doc.save(str(source));doc.close()
    store=lp_v3.Store(tmp_path/'data');store.ensure_dirs()
    event=lp_v3.new_event('Example');event['recipients']=[lp_v3.normalize_recipient({'name':'Example','to':'a@example.invalid','open_pw':'open','owner_pw':'owner'})]
    item=lp_v3.new_item('pdf');item['params'].update(files=[str(source)],watermark=False,encrypt=False);event['pipeline']=[item]
    assert store.save_event(event)[0]
    event,_=store.load_event(event['id'])
    return store,event


@pytest.mark.parametrize('mode',['mismatch','read_fail','accessor_after_save','set_fail'])
def test_saved_draft_encoding_warning_in_single_and_batch(tmp_path,monkeypatch,mode):
    import win32com.client
    saved=[]
    class Outlook:
        def CreateItem(self,n):
            mail=Mail(mode);saved.append(mail);return mail
    monkeypatch.setattr(win32com.client,'Dispatch',lambda _:Outlook())
    store,event=event_fixture(tmp_path)
    run=lp_v3.run_event(store,event,outdir=tmp_path/'out',word=object())
    assert saved[-1].saved==1 and run['stats']['ok']==1 and run['stats']['fail']==0
    assert run['items'][0]['encoding_warning'] and any('编码加固: 失败' in line for line in run['logs'])
    assert any('已保存' in warning for warning in run['warnings'])
    digest=hashlib.sha256(store.event_path(event['id']).read_bytes()).hexdigest()
    results=lp_v3.run_batch_events(store,[(event,digest)],tmp_path/'batch',object())
    assert results[0]['status']=='完成（需核对）' and '请勿重复生成' in results[0]['detail']
    assert saved[-1].saved==1 and len(saved)==2


def test_result_contains_per_property_readback_mismatch(monkeypatch):
    import win32com.client
    mail=Mail('mismatch')
    monkeypatch.setattr(win32com.client,'Dispatch',lambda _:type('O',(),{'CreateItem':lambda self,n:mail})())
    result={};lp_v3.create_outlook_draft({'to':'a@example.invalid'},[],'subject','body',hardening_result=result)
    assert mail.saved==1 and not result['ok'] and not result['persisted']
    assert any('message_codepage' in error and '不一致' in error for error in result['errors'])
