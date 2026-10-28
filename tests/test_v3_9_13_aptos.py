import tkinter as tk
from tkinter import ttk
import lp_tool, lp_v3
from test_v3_pool_gui import app,_event,_walk

def test_default_templates_and_roundtrip_keep_aptos_and_legacy_styles(tk_root):
    assert "font-family:Aptos," in lp_tool.DEFAULT_BODY
    assert "font-family:Aptos," in lp_v3.DEFAULT_BODY_V3
    root=tk_root
    try:
        a=tk.Text(root);b=tk.Text(root)
        sample='<div style="font-family:Aptos;font-size:14px;"><span style="font-size:12px">This email is sent on behalf of X.</span><p>Dear Investor,</p></div>'
        lp_tool.load_html(a,sample);lp_tool.load_html(b,sample)
        for widget in (a,b):
            style=lp_tool._style_at(widget,"1.0")
            assert style["f"]=="Aptos" and style["s"]==12
            tag=next(t for t in widget.tag_names("1.0") if t.startswith("st:"))
            assert widget.tag_cget(tag,"font")
            html=lp_tool.text_to_html(widget)
            assert "font-family:Aptos," in html and "font-size:12px" in html
            _,runs=lp_tool.html_to_runs(html)
            assert any("This email" in widget.get("1.0","end") and r[2].get("s")==12 for r in runs)
        legacy='<div style="font-family:Microsoft YaHei;font-size:14px">Old template</div>'
        lp_tool.load_html(a,legacy)
        assert "font-family:Microsoft YaHei" in lp_tool.text_to_html(a)
    finally:
        a.destroy(); b.destroy()

def test_toolbar_and_template_editor_default_to_aptos(app):
    e=_event(app,'Font UI',[{'name':'Test','to':'test@example.com'}])
    app.store.save_event(e);app.open_event(e['id']);app.root.deiconify()
    app._wb_show_step(2)
    for _ in range(5):app.root.update()
    combos=[w for w in _walk(app._wb_content) if isinstance(w,ttk.Combobox)]
    font=next(w for w in combos if "Aptos" in w.cget("values"))
    assert font.get()=="Aptos"
    assert str(font.cget("state"))=="readonly"
