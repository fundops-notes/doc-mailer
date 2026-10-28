import json
import lp_tool, lp_v3
from tkinter import ttk
from test_v3_pool_gui import app, no_messagebox, _event, _walk


def test_event_group_never_passes_through_pool(tmp_path):
    store=lp_v3.Store(tmp_path);store.ensure_dirs()
    original=lp_v3.normalize_recipient({'name':'Investor','group':'Old Event Group','to':'x@example.com'})
    ok,err,stats=store.merge_pool_records([original]);assert ok,err
    pool=store.list_pool();assert len(pool)==1 and pool[0]['group']==''
    # Legacy data remains stored for recovery; it never enters new events.
    f=store.pool_path(pool[0]['id']);raw=json.loads(f.read_text(encoding='utf-8'))
    raw['group']='Legacy Group';f.write_text(json.dumps(raw),encoding='utf-8')
    legacy=store.list_pool()[0];assert legacy['group']=='Legacy Group'
    event=lp_v3.new_event('New Event');updated,added,_=lp_v3.merge_recipients_into_event(event,[legacy])
    assert added==1 and updated['recipients'][0]['group']==''
    event['recipients']=[lp_v3.normalize_recipient({'name':'Investor','group':'Current Group'})]
    updated,_,_=lp_v3.merge_recipients_into_event(event,[legacy]);assert updated['recipients'][0]['group']=='Current Group'
    legacy['to']='new@example.com';ok,err,_=store.save_pool_recipient(legacy);assert ok,err
    assert json.loads(f.read_text(encoding='utf-8'))['group']=='Legacy Group'
    assert original['group']=='Old Event Group'


def test_pool_never_offers_event_group_but_event_editor_does(app):
    app._refresh_pool_tab();assert 'group' not in app._pool_tree['columns']
    pool=app._recipient_dialog({'name':'A','group':'Legacy'},[], 'pool',lambda out:True,allow_group=False)
    labels=[str(w.cget('text')) for w in _walk(pool) if isinstance(w,ttk.Label)]
    assert not any('合并发送组' in text for text in labels);pool.destroy()
    event=app._recipient_dialog({'name':'A','group':'Event'},[],'event',lambda out:True)
    labels=[str(w.cget('text')) for w in _walk(event) if isinstance(w,ttk.Label)]
    assert any('合并发送组' in text for text in labels);event.destroy()


def test_pool_excel_never_transfers_event_group():
    rows=lp_v3.pool_recipients_to_rows([{'name':'Investor','group':'Legacy','tags':['Fund']}])
    assert '合并组' not in rows[0]
    data=lp_v3.validate_pool_recipient_rows([['收件人名称','收件人','合并组'],['Investor','x@example.com','Old Event Group']])
    assert data['items'][0]['record']['group']==''
    event=lp_v3.validate_recipient_rows([['收件人名称','收件人','合并组'],['Investor','x@example.com','Current Event Group']])
    assert event['items'][0]['record']['group']=='Current Event Group'


def test_legacy_merge_group_preserved_on_edit(tmp_path):
    store=lp_v3.Store(tmp_path);store.ensure_dirs()
    r=lp_v3.new_pool_recipient({'name':'Legacy'})
    path=store.pool_path(r['id']);path.write_text(json.dumps({'id':r['id'],'name':'Legacy','merge_group':'Old Group'}),encoding='utf-8')
    loaded=store.list_pool()[0];assert loaded['group']=='Old Group'
    ok,err,_=store.save_pool_recipient(loaded);assert ok,err
    assert json.loads(path.read_text(encoding='utf-8'))['group']=='Old Group'
