"""Duplicate open passwords warn and do not stop processing."""
from pathlib import Path
import sys

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
import lp_v3
import lp_tool


def test_reused_open_password_warns_but_processing_succeeds(tmp_path):
    source=tmp_path/'report.pdf';source.write_bytes(b'%PDF fixture')
    event=lp_v3.new_event('warning test')
    event['recipients']=[lp_v3.normalize_recipient({'name':n,'to':n+'@example.com',
        'open_pw':'shared-open', 'owner_pw':'private-'+n}) for n in ('Alice','Bob')]
    item=lp_v3.new_item('pdf');item['params']['files']=[str(source)];event['pipeline']=[item]
    errors,warnings,_=lp_v3.preflight(event)
    assert not errors and any('打开口令重复' in w for w in warnings)
    assert not any('shared-open' in w for w in warnings)
    store=lp_v3.Store(tmp_path/'store');store.save_settings({'outdir':str(tmp_path/'out')})
    run=lp_v3.run_event(store,event,create_drafts=False,outdir=tmp_path/'out')
    assert run['status']=='ok' and run['stats']['ok']==2
    assert any('打开口令重复' in w for w in run['warnings'])


def test_owner_password_equal_to_open_is_still_blocked(tmp_path):
    source=tmp_path/'file.pdf';source.write_bytes(b'%PDF')
    event=lp_v3.new_event('owner guard');event['recipients']=[lp_v3.normalize_recipient({
        'name':'A','to':'a@example.com','open_pw':'same','owner_pw':'same'})]
    item=lp_v3.new_item('pdf');item['params']['files']=[str(source)];event['pipeline']=[item]
    errors,_,_=lp_v3.preflight(event)
    assert any('权限口令与打开口令相同' in x for x in errors)


def test_duplicate_password_warning_is_visible(monkeypatch):
    shown=[]
    monkeypatch.setattr(lp_tool.messagebox,'showwarning',lambda title,text:shown.append((title,text)))
    lp_tool.App._show_duplicate_password_warning(None,['其他提醒','打开口令重复（Alice、Bob）：同一口令可打开对方文件'])
    assert len(shown)==1 and '可以继续' in shown[0][1] and 'Alice' in shown[0][1]
