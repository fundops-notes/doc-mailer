import lp_v3
from test_v3_pool_gui import app,no_messagebox,_event


def test_event_table_mail_count_tracks_draft_units(app):
    e=_event(app,'Count test',[{'name':'A','to':'a@example.com','group':'G'},
         {'name':'B','to':'b@example.com','group':'G'}, {'name':'C','to':'c@example.com','group':''}])
    assert 'mail_count' in app.ev_tree['columns']
    for mode,expected in [('per_lp',3),('per_group',2),('all',1)]:
        e['options']['send_mode']=mode;ok,err=app.store.save_event(e);assert ok,err
        app._refresh_events_tab();row=app.ev_tree.item(e['id'],'values')
        assert int(row[app.ev_tree['columns'].index('mail_count')])==expected
        assert expected==len(lp_v3.draft_units(e))
    e['recipients']=[];e['options']['send_mode']='all';app.store.save_event(e);app._refresh_events_tab()
    row=app.ev_tree.item(e['id'],'values');assert int(row[app.ev_tree['columns'].index('mail_count')])==0
