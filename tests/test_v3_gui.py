# -*- coding: utf-8 -*-
"""v3 界面骨架冒烟测试：能建起来、能切页、能进事件工作台。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

tk = pytest.importorskip("tkinter")

import lp_v3
from lp_tool import App


@pytest.fixture()
def app(tmp_path, monkeypatch, tk_root):
    """在临时目录上建一个 App（不碰用户真实数据目录）。"""
    monkeypatch.setattr(lp_v3, "default_base_dir", lambda: tmp_path)
    a = App(tk_root)
    yield a
    a.shutdown()


@pytest.fixture(autouse=True)
def no_messagebox(monkeypatch):
    """测试里不弹模态框（模态框会挂住测试进程）。

    返回一个列表，记录所有被拦截的弹窗，测试可断言其内容。
    """
    import lp_tool
    calls = []

    class _MB:
        @staticmethod
        def showinfo(*a, **_k):
            calls.append(("info", a[0] if a else ""))

        @staticmethod
        def showwarning(*a, **_k):
            calls.append(("warning", a[0] if a else ""))

        @staticmethod
        def showerror(*a, **_k):
            calls.append(("error", a[0] if a else ""))

        @staticmethod
        def askyesno(*a, **_k):
            calls.append(("ask", a[0] if a else ""))
            return True

        @staticmethod
        def askokcancel(*a, **_k):
            calls.append(("ask", a[0] if a else ""))
            return True

    monkeypatch.setattr(lp_tool, "messagebox", _MB)
    return calls


def test_shell_and_all_screens(app):
    """v3.0.1：导航收敛成「事件库 / 设置 / 帮助」三个入口。"""
    for name in ("home", "settings", "help"):
        app.show_screen(name)
        assert name in app._screens


def test_nav_has_only_three_entries(app):
    """曾经的「场景与事件 / 场景模板 / 收件人库」三个入口已合并为事件库。"""
    app.show_screen("home")
    assert "scenes" not in app._screens
    assert "lib" not in app._screens


def test_event_library_has_event_and_template_tabs(app):
    """事件库的标签页：事件 / 场景模板 / 收件人库 / 执行历史。

    收件人库与执行历史（P4）做在**标签页**里，不占顶层导航入口 ——
    用户当初明确说过「三个入口看不懂区别，只要事件库」。
    """
    app.show_screen("home")
    tabs = [app.home_nb.tab(i, "text") for i in range(app.home_nb.index("end"))]
    assert len(tabs) == 4
    assert "事件" in tabs[0]
    assert "场景模板" in tabs[1]
    assert "收件人库" in tabs[2]
    assert "执行历史" in tabs[3]


def test_nav_still_has_only_three_entries(app):
    """P4 加了收件人库/执行历史，但**顶层导航仍然只有三项**。"""
    app.show_screen("home")
    assert "scenes" not in app._screens
    assert "lib" not in app._screens
    assert "pool" not in app._screens
    assert "runs" not in app._screens


def test_open_event_reopens_after_leaving_workbench(app):
    """回归：以前第二次点「打开工作台」毫无反应（页面被 pack_forget 了）。"""
    ev = lp_v3.new_event("重开事件")
    app.store.save_event(ev)
    app.open_event(ev["id"])
    assert app._screens["event"].winfo_manager() == "pack"
    app.show_screen("home")
    assert app._screens["event"].winfo_manager() == ""
    app.open_event(ev["id"])
    assert app._screens["event"].winfo_manager() == "pack"
    assert app._wb_event.get("id") == ev["id"]


def test_flow_chart_draws_event_pipeline(app):
    ev = lp_v3.new_event("流程图事件")
    ev["recipients"].append(lp_v3.normalize_recipient(
        {"name": "甲公司", "to": "a@x.com", "open_pw": "op1", "owner_pw": "ow1"}))
    ev["pipeline"] = [lp_v3.normalize_item({"kind": "generate",
                                            "params": {"template_path": "C:/t.docx"}}),
                      lp_v3.normalize_item({"kind": "attach",
                                            "params": {"files": ["C:/x.pdf"],
                                                       "encrypt": True}})]
    app.store.save_event(ev)
    app.show_screen("home")
    app._refresh_home()
    app.ev_tree.selection_set(ev["id"])
    app._draw_selected_flow()
    canvas = app.flow_canvas
    texts = [canvas.itemcget(i, "text") for i in canvas.find_all()
             if canvas.type(i) == "text"]
    joined = " ".join(t for t in texts if t)
    assert "收件人" in joined
    assert "加密" in joined
    assert "Outlook 草稿" in joined


def test_send_mode_switch_persists(app):
    ev = lp_v3.new_event("模式事件")
    ev["recipients"] = [
        lp_v3.normalize_recipient({"name": "甲", "to": "a@x.com",
                                   "open_pw": "op1", "owner_pw": "ow1"}),
        lp_v3.normalize_recipient({"name": "乙", "to": "b@x.com",
                                   "open_pw": "op2", "owner_pw": "ow2",
                                   "group": "G1"}),
    ]
    app.store.save_event(ev)
    app.open_event(ev["id"])
    app._wb_show_step(0)
    assert hasattr(app, "_wb_mode_var")
    app._wb_mode_var.set("每人一份")
    app._wb_change_send_mode()
    assert app._wb_event["options"]["send_mode"] == "per_lp"
    again, _err = app.store.load_event(ev["id"])
    assert again["options"]["send_mode"] == "per_lp"
    assert len(lp_v3.draft_units(again)) == 2


def test_create_scenario_and_event_from_home(app):
    app.show_screen("home")
    before = len(app.store.list_scenarios())
    sc = lp_v3.new_scenario("测试场景")
    sc["recipients"].append(lp_v3.normalize_recipient(
        {"name": "甲公司", "to": "a@x.com", "open_pw": "op1", "owner_pw": "ow1"}))
    assert app.store.save_scenario(sc)[0]
    app._refresh_home()
    assert len(app.store.list_scenarios()) == before + 1
    app._new_event(sc)
    evs = app.store.list_events()
    ev = next((e for e in evs if e.get("scenario_name") == "测试场景"), None)
    assert ev is not None, "新建的事件没找到：%s" % [e.get("name") for e in evs]
    assert ev["recipients"][0]["name"] == "甲公司"
    assert ev["scenario_name"] == "测试场景"


def test_workbench_all_five_steps(app):
    app.show_screen("home")
    ev = lp_v3.new_event("测试事件")
    ev["recipients"].append(lp_v3.normalize_recipient(
        {"name": "甲公司", "to": "a@x.com", "open_pw": "op1", "owner_pw": "ow1"}))
    app.store.save_event(ev)
    app.open_event(ev["id"])
    assert app._wb_event is not None
    for i in range(5):
        app._wb_show_step(i)
        assert app._wb_content.winfo_children(), "步骤 %d 没有内容" % (i + 1)
    # 步骤1 的收件人表里能看到刚加的人
    app._wb_show_step(0)
    tree = app._wb_rec_tree
    assert len(tree.get_children()) == 1
    assert tree.item(tree.get_children()[0])["values"][0] == "甲公司"


def test_event_info_edit_survives_save(app):
    ev = lp_v3.new_event("事件")
    app.store.save_event(ev)
    app.open_event(ev["id"])
    app._wb_event["name"] = "改过的名字"
    app._wb_save()
    again, err = app.store.load_event(ev["id"])
    assert err is None
    assert again["name"] == "改过的名字"


# ---------------------------------------------------------------- 流水线编辑

def _event_with(app, pipeline=None, recips=None):
    ev = lp_v3.new_event("事件")
    ev["recipients"] = [lp_v3.normalize_recipient(r) for r in (recips or [])]
    if pipeline is not None:
        ev["pipeline"] = [lp_v3.normalize_item(s) for s in pipeline]
    app.store.save_event(ev)
    app.open_event(ev["id"])
    return ev


def test_pipeline_item_editor_shows_table_and_deletes_selected(app):
    """v3.1：步骤卡 → 条目表格；Ctrl/Shift 多选后批量删除。"""
    ev = _event_with(app, pipeline=[{"kind": "attach",
                                     "params": {"files": ["C:/x.pdf"]}},
                                    {"kind": "attach",
                                     "params": {"files": ["C:/y.pdf"]}}])
    app._wb_show_step(1)
    assert len(app._wb_event["pipeline"]) == 2
    tree = app._wb_pipe_tree
    ids = list(tree.get_children())
    assert len(ids) == 2
    tree.selection_set(*ids)
    app._wb_del_selected_items()
    assert len(app._wb_event["pipeline"]) == 0


def test_pipeline_granularity_ui_removed_defaults_to_per_group(app):
    """v3.5.0：「产出颗粒度」设置已移除，默认按分组产出。"""
    ev = _event_with(app, pipeline=[{"kind": "attach",
                                     "params": {"files": ["C:/x.pdf"]}}])
    app._wb_show_step(1)
    item = app._wb_event["pipeline"][0]
    assert item["granularity"] == "per_group"
    tree = app._wb_pipe_tree
    assert "gran" not in list(tree["columns"])
    assert not any("产出颗粒度" in str(tree.heading(c)["text"])
                   for c in tree["columns"])
    assert not hasattr(app, "_wb_set_item_gran")


def test_pipeline_item_mapping_uses_field_level_source(app):
    """v3.1：来源下沉到字段级——占位符单独选来源。"""
    ev = _event_with(app, pipeline=[{"kind": "generate",
                                     "params": {"template_path": "C:/t.docx"}}],
                     recips=[{"name": "甲公司", "to": "a@x.com",
                              "open_pw": "op1", "owner_pw": "ow1",
                              "fields": {"Call_Amount": "200,000"}}])
    app._wb_show_step(1)
    item = app._wb_event["pipeline"][0]
    item["mapping"] = {"{{Call_Amount}}": "fields.Call_Amount"}
    opts = app._wb_field_source_options(item, app._wb_event["recipients"][0])
    assert ("fields.Call_Amount", "收件人字段", "Call_Amount") in opts


def test_pipeline_item_sheet_config_is_item_level(app):
    """外部数据表配置挂在条目上（不再是 mapping.sheet）。"""
    it = lp_v3.normalize_item({"kind": "generate",
                               "params": {"template_path": "C:/t.docx"},
                               "mapping": {"{{X}}": "sheet.X"},
                               "sheet": {"path": "C:/d.xlsx",
                                         "key_col": "Name"}})
    assert it["sheet"]["path"] == "C:/d.xlsx"
    assert it["sheet"]["key_col"] == "Name"
    assert it["mapping"]["{{X}}"] == "sheet.X"
    assert lp_v3.item_uses_sheet(it)


def test_detect_placeholders_button(app, word_template):
    ev = _event_with(app, pipeline=[{"kind": "generate",
                                     "params": {"template_path": str(word_template)}}],
                     recips=[{"name": "甲公司", "to": "a@x.com",
                              "open_pw": "op1", "owner_pw": "ow1",
                              "fields": {"Call_Amount": "200,000",
                                         "Due_Date": "2026-11-15"}}])
    item = app._wb_event["pipeline"][0]
    app._wb_detect_placeholders(item)
    mapped = item["mapping"]
    assert "{{LP_Name}}" in mapped
    assert mapped["{{Call_Amount}}"] == "fields.Call_Amount"
    assert mapped["{{Due_Date}}"] == "fields.Due_Date"


def test_pipeline_item_params_keep_tk_variables_alive(app):
    """回归：tk 变量被 GC 回收后 Entry 会变空（控件还在、值没了）。

    tk.StringVar 的 Python 侧一旦被回收，它的 __del__ 会把 Tcl 变量一起删掉。
    所以条目参数区里的变量必须被 App 吊住。
    """
    import gc
    ev = _event_with(app, pipeline=[{"kind": "generate",
                                     "name": "季度对账单",
                                     "params": {"template_path": "C:/t.docx",
                                                "to_pdf": True}}])
    app._wb_show_step(1)
    entries = []

    def walk(w):
        for c in w.winfo_children():
            if isinstance(c, tk.Entry):
                entries.append(c)
            walk(c)

    walk(app._wb_params_frame)
    assert entries, "参数区应该有输入框"
    name_entry = entries[0]
    assert name_entry.get() == "季度对账单", "名称输入框应显示条目名"
    gc.collect()
    assert name_entry.get() == "季度对账单", "GC 后名称输入框不应变空"


def test_run_panel_shows_preflight_errors(app):
    ev = _event_with(app, recips=[{"name": "甲", "to": "a@x.com",
                                   "open_pw": "", "owner_pw": ""}])
    app._wb_show_step(5)
    assert app._wb_content.winfo_children(), "预检面板应该有内容"


def test_run_panel_lists_expected_outputs(app, tmp_path):
    pdf = tmp_path / "src.pdf"
    import fitz
    doc = fitz.open()
    doc.new_page()
    doc.save(str(pdf))
    doc.close()
    ev = _event_with(app, pipeline=[{"kind": "attach",
                                     "params": {"files": [str(pdf)],
                                                "encrypt": True}}],
                     recips=[{"name": "甲公司", "to": "a@x.com",
                              "open_pw": "op12345678", "owner_pw": "ow12345678"}])
    app._wb_show_step(5)
    labels = []
    for c in app._wb_content.winfo_children():
        for sub in c.winfo_children():
            labels.append(sub)
    assert labels
