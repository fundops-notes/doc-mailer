import tkinter as tk
import lp_tool, lp_v3
from test_v3_pool_gui import app, no_messagebox, _event


def test_context_menu_targets_clicked_event_and_blank_is_safe(app, monkeypatch):
    first=_event(app, "First")
    second=_event(app, "Second")
    app._refresh_home(); app.root.deiconify()
    for _ in range(5): app.root.update()
    tree=app.ev_tree
    tree.selection_set(first["id"])
    box=tree.bbox(second["id"]); assert box
    shown=[]
    monkeypatch.setattr(app._event_context_menu, "tk_popup", lambda x,y: shown.append((x,y)))
    monkeypatch.setattr(app._event_context_menu, "grab_release", lambda: None)
    tree.event_generate("<Button-3>",x=box[0]+4,y=box[1]+box[3]//2)
    app.root.update()
    assert app._selected_event_id()==second["id"] and len(shown)==1
    menu=app._event_context_menu
    assert [menu.entrycget(i,"label") for i in range(5)]==["打开","编辑","复制","删除","导出"]
    tree.event_generate("<Button-3>",x=4,y=tree.winfo_height()-2);app.root.update()
    assert len(shown)==1 and app._selected_event_id()==second["id"]
    count_before=len(app.store.list_events())
    menu.invoke(2);app.root.update()
    copies=app.store.list_events()
    assert len(copies)==count_before+1 and any(e["name"].startswith("Second") and e["id"]!=second["id"] for e in copies)
    tree.selection_set(second["id"]);menu.invoke(3);app.root.update()
    assert app.store.load_event(second["id"])[0] is None
    assert app.store.load_event(first["id"])[0] is not None


def test_menu_dispatches_all_five_actions_to_clicked_record(tmp_path,monkeypatch,tk_root):
    monkeypatch.setattr(lp_v3,'default_base_dir',lambda:tmp_path)
    monkeypatch.setattr(lp_tool.App,'check_update',lambda *a,**k:None)
    calls=[]
    methods=['_open_selected_event','_edit_event_dialog','_dup_event','_del_event','_lib_export_event']
    for name in methods:
        monkeypatch.setattr(lp_tool.App,name,lambda self,n=name:calls.append((n,self._selected_event_id())))
    a=lp_tool.App(tk_root)
    try:
        ev=_event(a,'Context target');a._refresh_home();a.root.deiconify()
        for _ in range(4):a.root.update()
        a.ev_tree.selection_set(ev['id'])
        for i in range(5):a._event_context_menu.invoke(i)
        assert calls==[(name,ev['id']) for name in methods]
    finally:
        a.shutdown()
