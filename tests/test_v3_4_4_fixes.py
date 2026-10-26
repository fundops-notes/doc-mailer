# -*- coding: utf-8 -*-
"""v3.4.4 修复回归测试（对应验收报告里的 P0/P1/P2）。

覆盖：
  · 收件人表列数与 values 对齐、权限口令遮罩、自定义字段不错位
  · 编辑收件人保存保留旧 rec["attachments"]
  · 导入模板不含「附件路径」；导入该列只进自定义字段
  · 流水线页左右分栏：左栏=条目表+参数区，右栏=统一附件+处理+对应表（且真正可见）
  · 流水线条目表横向滚动条
  · Word-MailMerge 参数区「导出要素模板」「下载示例 Word 模板」按钮
  · 附加文件「导出路径模板」
  · 「发送人统一密码」落盘并同步到 event.options（引擎可读）
  · 帮助页旧文案已清理
  · 重复内容提示三处统一为 [附件重复提示] … 请人工确认
  · process_overrides 多文件存列表
  · 版本号 3.4.4
"""
import sys
import traceback
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

tk = pytest.importorskip("tkinter")

import lp_v3
import lp_tool
from lp_tool import App


def _mini_docx(path, text):
    """最小 .docx（不依赖 Word/python-docx）：占位符检测只读 zip 里的 document.xml。"""
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body><w:p><w:r><w:t>%s</w:t></w:r></w:p></w:body></w:document>' % text)
    ct = ('<?xml version="1.0" encoding="UTF-8"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-'
          'officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    with zipfile.ZipFile(str(path), "w") as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)


class _MB:
    def __init__(self):
        self.calls = []
        self.answer3 = True

    def showinfo(self, title, msg="", **k):
        self.calls.append(("info", str(title), str(msg)))

    def showwarning(self, title, msg="", **k):
        self.calls.append(("warning", str(title), str(msg)))

    def showerror(self, title, msg="", **k):
        self.calls.append(("error", str(title), str(msg)))

    def askyesno(self, title, msg="", **k):
        self.calls.append(("ask", str(title), str(msg)))
        return True

    def askokcancel(self, title, msg="", **k):
        self.calls.append(("confirm_save", str(title), str(msg)))
        return bool(self.answer3)

    def askyesnocancel(self, title, msg="", **k):
        self.calls.append(("ask3", str(title), str(msg)))
        return self.answer3


class _FD:
    save_path = ""
    open_path = ""
    dir_path = ""

    def asksaveasfilename(self, **k):
        return self.save_path

    def askopenfilename(self, **k):
        return self.open_path

    def askdirectory(self, **k):
        return self.dir_path


@pytest.fixture()
def app(tmp_path, monkeypatch, tk_root):
    monkeypatch.setattr(lp_v3, "default_base_dir", lambda: tmp_path)
    mb = _MB()
    fd = _FD()
    monkeypatch.setattr(lp_tool, "messagebox", mb)
    monkeypatch.setattr(lp_tool, "filedialog", fd)
    monkeypatch.setattr(lp_tool.App, "check_update", lambda self, *a, **k: None)
    errors = []
    old_hook = tk_root.report_callback_exception
    tk_root.report_callback_exception = lambda *a: errors.append(
        "".join(traceback.format_exception(*a)))
    a = App(tk_root)
    tk_root.update()
    try:
        yield a, mb, fd, errors
    finally:
        a.shutdown()
        tk_root.report_callback_exception = old_hook


def _walk(w):
    yield w
    for c in w.winfo_children():
        yield from _walk(c)


def _cls(w):
    try:
        return w.winfo_class()
    except Exception:
        return "?"


def _buttons(where, prefix=None):
    out = []
    for w in _walk(where):
        if _cls(w) == "TButton":
            t = str(w.cget("text"))
            if prefix is None or t.startswith(prefix):
                out.append((t, w))
    return out


def _find_toplevel(root, title):
    for w in root.winfo_children():
        try:
            if isinstance(w, tk.Toplevel) and w.winfo_exists() and w.title() == title:
                return w
        except Exception:
            pass
    return None


def _descendant(w, anc):
    cur = w
    while cur is not None:
        if cur is anc:
            return True
        cur = getattr(cur, "master", None)
    return False


def _save_event(app, name, recips, pipeline=None):
    ev = lp_v3.new_event(name)
    ev["recipients"] = [lp_v3.normalize_recipient(r) for r in recips]
    if pipeline:
        ev["pipeline"] = [lp_v3.normalize_item(s) for s in pipeline]
    app.store.save_event(ev)
    return ev


def test_version_is_381():
    assert lp_tool.APP_VERSION == "3.9.2"
    assert lp_v3.APP_VERSION == "3.9.2"


def test_recipient_table_values_match_columns_and_mask(app, tmp_path):
    a, _mb, _fd, _err = app
    att = tmp_path / "专属.pdf"
    att.write_bytes(b"%PDF-1.4 x")
    ev = _save_event(a, "列对齐", [{
        "name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
        "owner_pw": "Owner123456", "attachments": [str(att)],
        "fields": {"Call_Amount": "200,000", "Due_Date": "2026-11-15"},
    }])
    a.open_event(ev["id"])
    a._wb_show_step(0)
    a.root.update()
    tree = a._wb_rec_tree
    cols = list(tree["columns"])
    vals = list(tree.item(tree.get_children()[0])["values"])
    assert len(vals) == len(cols), "values 必须与 columns 一一对应：%s vs %s" % (vals, cols)
    heads = [tree.heading(c)["text"] for c in cols]
    assert "附件路径" not in heads
    # v3.7.2：自定义字段不在收件人信息里显示
    assert "Call_Amount" not in heads and "Due_Date" not in heads
    assert vals[cols.index("owner_pw")] == "••••", "权限口令必须打码"
    assert vals[cols.index("open_pw")] == "Open123456"
    # 记录里的 attachments 本体仍在（流水线兼容旧数据）
    saved, _e = a.store.load_event(ev["id"])
    assert saved["recipients"][0]["attachments"] == [str(att)]


def test_edit_recipient_keeps_legacy_attachments(app, tmp_path):
    a, _mb, _fd, _err = app
    att = tmp_path / "旧附件.pdf"
    att.write_bytes(b"%PDF-1.4 x")
    ev = _save_event(a, "编辑保留附件", [{
        "name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
        "owner_pw": "Owner123456", "attachments": [str(att)],
    }])
    a.open_event(ev["id"])
    a._wb_show_step(0)
    a.root.update()
    tree = a._wb_rec_tree
    tree.selection_set(tree.get_children()[0])
    a._wb_edit_recipient()
    a.root.update()
    dlg = _find_toplevel(a.root, "编辑收件人")
    assert dlg is not None
    for t, b in _buttons(dlg):
        if t == "保存":
            b.invoke()
            break
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["recipients"][0]["attachments"] == [str(att)], \
        "编辑收件人不能清空旧 rec['attachments']"


def test_import_template_has_no_attachment_column(app, tmp_path):
    a, _mb, fd, _err = app
    out = tmp_path / "导入模板.xlsx"
    fd.save_path = str(out)
    a._wb_export_template()
    print(out)
    import openpyxl
    ws = openpyxl.load_workbook(str(out))["收件人清单"]
    headers = [c.value for c in ws[1]]
    assert "附件路径" not in headers, "导入模板不能再有「附件路径」列：%s" % headers
    assert headers == lp_v3.import_template_headers(["Call_Amount", "Unfunded", "Due_Date"])


def test_import_attachment_column_becomes_custom_field():
    rows = [["收件人名称", "收件人", "附件路径"],
            ["甲公司", "a@x.com", r"C:\a.pdf"]]
    recips, unknown, field_cols = lp_v3.parse_import_rows_v3(rows)
    rec = lp_v3.normalize_recipient(recips[0])
    assert rec["attachments"] == []
    assert rec["fields"]["附件路径"] == r"C:\a.pdf"
    assert "附件路径" in unknown and "附件路径" in field_cols


def test_pipeline_left_right_layout_and_scrollbar(app, tmp_path):
    a, _mb, _fd, _err = app
    tpl = tmp_path / "t.docx"
    _mini_docx(tpl, "Hello {{Client_Name}}")
    att = tmp_path / "a.pdf"
    att.write_bytes(b"%PDF-1.4 x")
    ev = _save_event(a, "分栏", [{"name": "甲", "to": "a@x.com",
                                  "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "generate", "params": {"template_path": str(tpl)}},
                               {"kind": "attach", "params": {"files": [str(att)]}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()

    paned = None
    for w in _walk(a._wb_content):
        if _cls(w) == "TPanedwindow":
            paned = w
            break
    assert paned is not None, "流水线页应有 PanedWindow"
    assert str(paned.cget("orient")) == "horizontal"
    panes = [paned.nametowidget(x) for x in paned.panes()]
    assert len(panes) == 2
    left, right = panes
    assert len(left.winfo_children()) >= 2, "左栏必须有条目表 + 参数区"
    assert _descendant(a._wb_pipe_tree, left), "条目表必须在左栏"
    assert _descendant(a._wb_params_frame, left), "参数区必须在左栏"
    assert _descendant(a._wb_map_tree, right), "对应表必须在右栏"
    assert a._wb_process_btn is None, "处理动作合并到下一步"
    assert not _buttons(a._wb_content, "📎 统一附件")
    assert a._wb_pipe_tree.winfo_manager() == "grid", "条目表和完整宽度的滚动条用grid共同布局，且仍属于左栏"
    # ⑥ 横向滚动条
    assert str(a._wb_pipe_tree.cget("xscrollcommand")), "流水线条目表必须有横向滚动条"
    horiz = [w for w in _walk(a._wb_pipe_tree.master)
             if _cls(w) == "TScrollbar" and str(w.cget("orient")) == "horizontal"]
    assert horiz, "流水线条目表容器里应有横向滚动条控件"
    assert "处理" in a._wb_next_btn.cget("text")


def test_word_mailmerge_missing_buttons_are_wired(app, tmp_path):
    a, _mb, fd, _err = app
    tpl = tmp_path / "t.docx"
    _mini_docx(tpl, "Dear {{Client_Name}} {{Call_Amount}}")
    ev = _save_event(a, "按钮", [{"name": "甲", "to": "a@x.com",
                                  "open_pw": "Open123456", "owner_pw": "Owner123456",
                                  "fields": {"Call_Amount": "100"}}],
                     pipeline=[{"kind": "generate", "params": {"template_path": str(tpl)}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    texts = [t for t, _b in _buttons(a._wb_content)]
    assert any("导出要素模板" in t for t in texts), texts
    assert any("下载示例 Word 模板" in t for t in texts), texts
    # 元素模板导出：行=收件人、列=检测出的占位符
    out = tmp_path / "要素模板.xlsx"
    fd.save_path = str(out)
    _buttons(a._wb_content, "导出要素模板")[0][1].invoke()
    a.root.update()
    import openpyxl
    ws = openpyxl.load_workbook(str(out))["要素模板"]
    headers = list([c.value for c in ws[1]])
    assert headers[0] == "收件人名称"
    assert set(headers[1:]) == {"{{Client_Name}}", "{{Call_Amount}}"}, headers
    assert [c.value for c in ws[2]][0] == "甲"


def test_attach_item_has_path_template_export(app, tmp_path):
    a, _mb, fd, _err = app
    att = tmp_path / "a.pdf"
    att.write_bytes(b"%PDF-1.4 x")
    ev = _save_event(a, "路径模板", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": [], "input_type": "excel", "path_source": "item.paths", "path_map": {"甲": [str(att)]}}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    tree = a._wb_pipe_tree
    tree.selection_set(tree.get_children()[0])
    a.root.update()
    assert any("导出对应模板" in t for t, _b in _buttons(a._wb_content))
    out = tmp_path / "路径模板.xlsx"
    fd.save_path = str(out)
    a._wb_export_attach_path_template(a._wb_event["pipeline"][0])
    import openpyxl
    ws = openpyxl.load_workbook(str(out))["附件路径"]
    rows = [[c.value for c in r] for r in ws.iter_rows()]
    assert rows[0] == ["收件人名称", "附件路径"]
    assert rows[1][0] == "甲" and rows[1][1] == str(att)


def test_gp_unified_password_persists_and_reaches_engine(app):
    a, _mb, _fd, errors = app
    ev = _save_event(a, "统一密码", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "generate",
                                "params": {"to_pdf": True, "watermark": True,
                                           "encrypt": True}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    gp = None
    for w in _walk(a._wb_params_frame):
        if _cls(w) == "TCheckbutton" and "发送人统一密码" in str(w.cget("text")):
            gp = w
            break
    assert gp is not None, "勾了加密之后应出现「发送人统一密码」"
    gp.invoke()        # 勾上（回调里落盘 + 同步 event.options）
    a.root.update()
    entry = None
    for w in _walk(gp.master):
        if _cls(w) == "TEntry":
            entry = w
            break
    assert entry is not None
    entry.delete(0, "end")
    entry.insert(0, "GPsecret123")
    entry.event_generate("<KeyRelease>", keysym="a")
    a.root.update()
    assert not errors, "勾选统一密码不应再抛异常：%s" % errors[-1]
    # v3.5.1：全局 GP 设置已删除，条目级开关只写 event.options
    saved, _e2 = a.store.load_event(ev["id"])
    assert saved["options"]["gp_owner_mode"] is True
    assert saved["options"]["gp_owner_pw"] == "GPsecret123", "引擎只读 event.options，必须同步"


def test_help_page_text_updated(app):
    a, _mb, _fd, _err = app
    a.show_screen("help")
    a.root.update()
    texts = []
    for w in _walk(a._screens["help"]):
        if _cls(w) == "TLabel":
            texts.append(str(w.cget("text")))
    joined = "\n".join(texts)
    # v3.5.0：颗粒度设置已移除，帮助页改为说明「按分组」
    assert "预生成" in joined
    assert "生成全部 Outlook 草稿" in joined
    assert "预览与生成" not in joined
    for stale in ("产出粒度", "路径来源", "生成文档"):
        assert stale not in joined, "帮助页仍有旧文案：%s" % stale


def test_duplicate_prompt_wording_unified(tmp_path):
    f1 = tmp_path / "A.pdf"
    f2 = tmp_path / "B.pdf"
    payload = b"%PDF-1.4 same-bytes"
    f1.write_bytes(payload)
    f2.write_bytes(payload)
    ev = lp_v3.new_event("重复")
    ev["recipients"] = [lp_v3.normalize_recipient({
        "name": "甲", "to": "a@x.com", "open_pw": "Open123456",
        "owner_pw": "Owner123456"})]
    ev["attachments"] = [
        lp_v3.normalize_attachment({"name": "A", "path": str(f1), "scope": "all"}),
        lp_v3.normalize_attachment({"name": "B", "path": str(f2), "scope": "all"}),
    ]
    _errors, warnings, _info = lp_v3.preflight(ev)
    hit = [w for w in warnings if "重复" in w]
    assert hit
    assert "[附件重复提示]" in hit[0] and "请人工确认" in hit[0], hit[0]


def test_process_overrides_multi_file_saved_as_list(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "overrides", [{"name": "甲", "to": "a@x.com",
                                       "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": ["C:/a.pdf"]}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    mt = a._wb_map_tree
    mt.delete(*mt.get_children())
    mt.insert("", "end", values=["甲", "a.pdf\nb.pdf", "C:/a.pdf / C:/b.pdf", "已处理"])
    a._wb_map_save_edits()
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["process_overrides"]["甲"] == ["C:/a.pdf", "C:/b.pdf"]


# ---------------------------------------------------------------- v3.4.6

def test_item_rename_direct_without_dialog(app):
    """名称输入框 +「改名」直接生效，不再弹重命名对话框。"""
    a, _mb, _fd, _err = app
    ev = _save_event(a, "改名", [{"name": "甲", "to": "a@x.com",
                                  "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "name": "旧名字",
                                "params": {"files": ["C:/a.pdf"]}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    assert not hasattr(a, "_wb_rename_item"), "重命名对话框应已删除"
    entry = a._wb_item_name_entry
    entry.delete(0, "end")
    entry.insert(0, "新名字")
    for t, b in _buttons(a._wb_params_frame):
        if t == "改名":
            b.invoke()
            break
    a.root.update()
    item = a._wb_event["pipeline"][0]
    assert item["name"] == "新名字"
    saved, _e = a.store.load_event(ev["id"])
    assert saved["pipeline"][0]["name"] == "新名字", "改名必须立即持久化"
    assert _find_toplevel(a.root, "重命名条目") is None, "不应再弹出重命名对话框"


def test_double_click_item_focuses_name_entry_without_dialog(app):
    """双击条目行只是聚焦名称输入框，不弹窗。"""
    a, _mb, _fd, _err = app
    ev = _save_event(a, "双击", [{"name": "甲", "to": "a@x.com",
                                  "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": ["C:/a.pdf"]}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    tree = a._wb_pipe_tree
    iid = tree.get_children()[0]
    tree.identify_column = lambda _x: "#2"          # 模拟双击「条目名称」列
    tree.identify_row = lambda _y: iid
    import types
    a._wb_on_pipe_double(types.SimpleNamespace(x=10, y=10))
    a.root.update()
    assert _find_toplevel(a.root, "重命名条目") is None
    assert a._wb_item_name_entry is not None


def test_import_field_template_applies_values_and_mapping(app, tmp_path):
    """导出要素模板后必须能导回来：按收件人名称匹配、写入 fields、自动接映射。"""
    a, mb, fd, _err = app
    tpl = tmp_path / "t.docx"
    _mini_docx(tpl, "Dear {{Client_Name}} Amount {{Call_Amount}}")
    ev = _save_event(a, "要素导入", [
        {"name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
         "owner_pw": "Owner123456", "fields": {"Call_Amount": "OLD"}},
        {"name": "乙公司", "to": "b@x.com", "open_pw": "Open223456",
         "owner_pw": "Owner223456"},
    ], pipeline=[{"kind": "generate", "params": {"template_path": str(tpl)}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    item = a._wb_event["pipeline"][0]
    # 造一份「要素模板」：收件人名称 + 占位符列（与导出的格式一致）
    import openpyxl
    src = tmp_path / "要素.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["收件人名称", "{{Call_Amount}}", "{{Client_Name}}", "{{Due_Date}}"])
    ws.append(["甲公司", 200000, "甲方科技", 46193])
    ws.append(["乙公司", 300000, "乙方科技", 46042])
    ws.append(["查无此人", 999, "X", "2027-01-01"])
    wb.save(str(src))
    fd.open_path = str(src)
    a._date_import_mode = lambda cols, count: "iso"
    a._wb_import_field_template(item)
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    r1 = saved["recipients"][0]
    r2 = saved["recipients"][1]
    assert r1["fields"]["Call_Amount"] == "200000"
    assert r1["fields"]["Client_Name"] == "甲方科技"
    assert r2["fields"]["Call_Amount"] == "300000"
    # 46193 / 46042 是 General 格式的 Excel 日期序列号，应按日期导入
    assert r1["fields"]["Due_Date"] == "2026-06-20"
    assert r2["fields"]["Due_Date"] == "2026-01-20"
    assert saved["pipeline"][0]["mapping"]["{{Call_Amount}}"] == "fields.Call_Amount"
    assert saved["pipeline"][0]["mapping"]["{{Client_Name}}"] == "fields.Client_Name"
    msgs = [m for m in mb.calls if m[0] == "info" and "已导入" in m[1]]
    assert msgs, "导入后应有汇总提示"
    assert "查无此人" in msgs[-1][2], "未匹配的收件人应在提示里列出"


def test_mailmerge_uses_inline_hint_not_help_dialog(app, tmp_path):
    """v3.4.6：不需要「Word-MailMerge 说明」按钮/弹窗，关键信息直接显示。"""
    a, _mb, _fd, _err = app
    tpl = tmp_path / "t.docx"
    _mini_docx(tpl, "{{Client_Name}}")
    ev = _save_event(a, "内联说明", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456",
                                      "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "generate", "params": {"template_path": str(tpl)}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    texts = [t for t, _b in _buttons(a._wb_content)]
    assert not any("说明" in t for t in texts), "不应再有说明按钮：%s" % texts
    assert not hasattr(a, "_show_mailmerge_help"), "说明弹窗方法应删除"
    labels = [str(w.cget("text")) for w in _walk(a._wb_params_frame)
              if _cls(w) == "TLabel"]
    joined = "\n".join(labels)
    assert "占位符写法" in joined
    assert "转 PDF → 加水印 → 加密" in joined
    assert "水印文字与各人打开口令在「收件人」处维护" in joined
    assert any("导入要素模板" in t for t, _b in _buttons(a._wb_content))


# ---------------------------------------------------------------- v3.4.7 检查更新

class _FakeResp:
    def __init__(self, url="", body=b"{}"):
        self._url = url
        self._body = body

    def geturl(self):
        return self._url

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def test_check_latest_web_fallback_when_api_rate_limited(monkeypatch):
    """APP 以前的唯一路径 api.github.com 常被 403 限流；现在网页跳转仍能拿到版本。"""
    import urllib.error
    calls = []

    def fake_urlopen(req, timeout=None, **k):
        url = req.full_url if hasattr(req, "full_url") else str(req)
        calls.append(url)
        if "api.github.com" in url:
            raise urllib.error.HTTPError(url, 403, "rate limit exceeded",
                                         {"X-RateLimit-Remaining": "0"}, None)
        return _FakeResp("https://github.com/fundops-notes/doc-mailer/releases/tag/v9.9.9")

    import urllib.request
    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    a = object.__new__(lp_tool.App)
    a._last_update_err = ""
    info = lp_tool.App.check_latest(a, timeout=5)
    assert info is not None
    assert info["tag"] == "v9.9.9"
    assert info["asset_url"].endswith("/download/v9.9.9/doc-mailer_v9.9.9_portable.zip")
    assert "api.github.com" in calls[0] or "api.github.com" in calls[-1]
    assert "rate limit" in a._last_update_err


def test_check_latest_api_still_supplies_notes_and_asset(monkeypatch):
    import json as _json
    body = _json.dumps({"tag_name": "v9.9.9", "body": "更新说明",
                        "html_url": "https://github.com/fundops-notes/doc-mailer/releases/tag/v9.9.9",
                        "assets": [{"name": "doc-mailer_v9.9.9_portable.zip",
                                    "browser_download_url": "https://example.com/a.zip",
                                    "size": 123, "digest": "sha256:abc"}]}).encode("utf-8")

    def fake_urlopen(req, timeout=None, **k):
        url = req.full_url if hasattr(req, "full_url") else str(req)
        if "api.github.com" in url:
            return _FakeResp(url, body)
        return _FakeResp("https://github.com/fundops-notes/doc-mailer/releases/tag/v9.9.9")

    import urllib.request
    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    a = object.__new__(lp_tool.App)
    a._last_update_err = ""
    info = lp_tool.App.check_latest(a, timeout=5)
    assert info["body"] == "更新说明"
    assert info["asset_url"] == "https://example.com/a.zip"
    assert info["asset_size"] == 123
    assert info["asset_digest"] == "sha256:abc"


def test_check_latest_returns_none_when_both_fail(monkeypatch):
    def fake_urlopen(req, timeout=None, **k):
        raise OSError("network down")

    import urllib.request
    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    a = object.__new__(lp_tool.App)
    a._last_update_err = ""
    assert lp_tool.App.check_latest(a, timeout=1) is None
    assert "网页跳转" in a._last_update_err
    assert "GitHub API" in a._last_update_err

def test_ver_tuple_is_static_and_show_update_works(app):
    """回归：_ver_tuple 曾缺 @staticmethod，_show_update 里 self._ver_tuple(...) 抛
    TypeError，被 _poll_queue 吞掉 → 点「检查更新」永远没反应。"""
    a, mb, _fd, _err = app
    assert lp_tool.App._ver_tuple("v3.4.7") == (3, 4, 7)
    info = {"tag": "v3.4.7", "body": "", "url": "u",
            "asset_url": None, "asset_size": None, "asset_digest": None}
    a._show_update(info, False)     # 不应抛异常
    hits = [m for m in mb.calls if m[1] == "检查更新"]
    assert hits and "已是最新版本" in hits[-1][2]

def test_mapping_dialog_lists_all_placeholders_and_preselects_row(app, tmp_path):
    """回归：映射弹窗的占位符下拉曾漏传 values（只显示第一个），且双击哪行都编辑第一个。"""
    a, _mb, _fd, _err = app
    tpl = tmp_path / "t.docx"
    _mini_docx(tpl, "{{Client_Name}} {{Call_Amount}} {{Due_Date}}")
    ev = _save_event(a, "映射弹窗", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456",
                                      "fields": {"Call_Amount": "100"}}],
                     pipeline=[{"kind": "generate", "params": {"template_path": str(tpl)}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    item = a._wb_event["pipeline"][0]
    item["mapping"] = {"{{Client_Name}}": "name",
                       "{{Call_Amount}}": "fields.Call_Amount",
                       "{{Due_Date}}": "date"}
    a._wb_render_params(item)
    a.root.update()
    # 指定第二个占位符打开弹窗（模拟双击第二行）
    a._wb_edit_mapping(item, "{{Call_Amount}}")
    a.root.update()
    dlg = _find_toplevel(a.root, "占位符来源")
    assert dlg is not None
    boxes = [w for w in _walk(dlg) if _cls(w) == "TCombobox"]
    assert len(boxes) == 2
    ph_box, src_box = boxes
    values = list(ph_box.cget("values"))
    assert set(values) == set(item["mapping"].keys()), values
    assert len(values) == 3
    assert ph_box.get() == "{{Call_Amount}}", "应按双击行预选对应占位符"
    assert "本事件字段：Call_Amount" in src_box.get()
    # 切换到第三个占位符 → 来源下拉要联动
    ph_box.set("{{Due_Date}}")
    ph_box.event_generate("<<ComboboxSelected>>")
    a.root.update()
    assert "日期" in src_box.get(), "切换占位符后来源下拉应跟着刷新，实际=%r" % src_box.get()
    # 选一个来源并确定
    want = [v for v in src_box.cget("values") if "本事件字段：Call_Amount" in v][0]
    src_box.set(want)
    for t, b in _buttons(dlg):
        if t == "确定":
            b.invoke()
            break
    a.root.update()
    assert item["mapping"]["{{Due_Date}}"] == "fields.Call_Amount"
    saved, _e = a.store.load_event(ev["id"])
    assert saved["pipeline"][0]["mapping"]["{{Due_Date}}"] == "fields.Call_Amount"

def test_self_update_progress_hook_and_staging(app, tmp_path, monkeypatch):
    """回归：自动更新下载时 _prog 只收 1 个参数，urlretrieve 传 3 个 → TypeError。
    这里用假 urlretrieve 走完整下载→解压→生成更新脚本链路（不真重启）。"""
    import queue
    import shutil
    import time
    import urllib.request
    import zipfile

    a, _mb, _fd, _err = app
    appdir = tmp_path / "appdir"
    appdir.mkdir()
    monkeypatch.setattr(lp_tool, "app_dir", lambda: appdir)

    fake_zip = tmp_path / "fake_update.zip"
    with zipfile.ZipFile(str(fake_zip), "w") as z:
        z.writestr("app/lp_tool.py", 'APP_VERSION = "9.9.9"\n')
        z.writestr("app/lp_v3.py", 'APP_VERSION = "9.9.9"\n')
        z.writestr("app/watermark_encrypt.py", "# test payload\n")
        z.writestr("app/start.bat", "@echo off\n")

    def fake_retrieve(url, filename, reporthook=None, data=None):
        # urlretrieve 的 reporthook 固定按 3 个参数调用
        if reporthook:
            reporthook(1, 1024, 4096)
            reporthook(4, 1024, 4096)
        shutil.copyfile(str(fake_zip), filename)
        return str(filename), {}

    monkeypatch.setattr(urllib.request, "urlretrieve", fake_retrieve)
    info = {"asset_url": "http://example.invalid/doc-mailer_portable.zip",
            "asset_digest": "", "tag": "v9.9.9",
            "asset_size": fake_zip.stat().st_size}
    a._self_update(info)
    items = []
    end = time.time() + 10
    while time.time() < end:
        try:
            while True:
                items.append(a._q.get_nowait())
        except queue.Empty:
            pass
        if any(k == "update_ready" for k, _ in items):
            break
        time.sleep(0.05)
    errs = [i for i in items if i[0] == "update_err"]
    assert not errs, "自动更新不应失败：%s" % errs
    assert any(k == "update_ready" for k, _ in items), items
    assert (appdir / ".update_staging" / "app" / "lp_tool.py").exists()
    bat = (appdir / "run_update.bat").read_text(encoding="ascii", errors="ignore")
    assert "run_update.ps1" in bat
    helper = (appdir / "run_update.ps1").read_text(encoding="utf-8-sig")
    assert "lp_data" in helper and "Copy-Item -LiteralPath" in helper and "Start-Process" in helper
    assert not (appdir / ".update_pkg.zip").exists(), "下载包应在解压后删除"

# ---------------------------------------------------------------- v3.5.0

def test_read_spreadsheet_uses_excel_display_values(tmp_path):
    """Excel 日期/金额/百分比按显示值读取，而不是变成序列号或 datetime 字符串。"""
    import datetime
    import openpyxl
    path = tmp_path / "显示格式.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "日期"
    ws["A2"] = datetime.datetime(2026, 11, 15)
    ws["A2"].number_format = 'yyyy"年"m"月"d"日"'
    ws["B1"] = "金额"
    ws["B2"] = 200000
    ws["B2"].number_format = "#,##0"
    ws["C1"] = "百分比"
    ws["C2"] = 0.256
    ws["C2"].number_format = "0.00%"
    wb.save(str(path))
    rows = lp_v3.read_spreadsheet(str(path))
    assert rows[1][0] == "2026年11月15日"
    assert rows[1][1] == "200,000"
    assert rows[1][2] == "25.6%"


def test_attach_excel_import_maps_paths_per_recipient(app, tmp_path, monkeypatch):
    """「从 Excel 导入路径」必须按收件人对应，不能把全部路径发给所有人。"""
    import openpyxl
    a, _mb, _fd, _err = app
    ev = _save_event(a, "路径导入", [
        {"name": "甲", "to": "a@x.com", "open_pw": "Open123456", "owner_pw": "Owner123456"},
        {"name": "乙", "to": "b@x.com", "open_pw": "Open223456", "owner_pw": "Owner223456"},
    ], pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    xlsx = tmp_path / "路径模板.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["收件人名称", "附件路径"])
    ws.append(["甲", str(tmp_path / "甲.pdf")])
    ws.append(["乙", str(tmp_path / "乙.pdf")])
    wb.save(str(xlsx))
    _fd.open_path = str(xlsx)
    a._wb_import_attach_paths(a._wb_event["pipeline"][0], None)
    a.root.update()
    dlg = _find_toplevel(a.root, "从 Excel 导入路径")
    assert dlg is not None
    for t, b in _buttons(dlg):
        if t == "导入":
            b.invoke()
            break
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    recs = {r["name"]: r for r in saved["recipients"]}
    assert not recs["甲"]["attachments"]
    assert not recs["乙"]["attachments"]
    it = saved["pipeline"][0]
    assert it["params"]["path_source"] == "item.paths"
    assert it["params"]["files"] == []
    # 引擎按人取路径
    unit = {"rep": recs["甲"], "recipients": [recs["甲"]], "lp_name": "甲",
            "group_key": "甲", "group": ""}
    paths, _label = lp_v3.resolve_item_paths(it, unit, saved)
    assert paths == [str(tmp_path / "甲.pdf")]
    # 预检不再误报「没有选择任何文件」
    errors, _w, _i = lp_v3.preflight(saved)
    assert not any("没有选择任何文件" in e for e in errors)


def test_left_nav_and_saas_styles(app):
    """v3.8.0：左侧导航 + Modern Enterprise SaaS 主题。"""
    a, _mb, _fd, _err = app
    a.root.update()
    assert set(a._nav_buttons.keys()) == {"events", "pool", "runs", "settings", "help"}
    for b in a._nav_buttons.values():
        assert not _descendant(b, a.content), "导航按钮应在内容区外的左栏"
    assert a._nav_buttons["events"].cget("style") in ("NavOn.TButton", "Nav.TButton")
    # 工作台三个入口并列，点击切换内容且高亮
    a.show_home_section("pool")
    a.root.update()
    assert a.home_nb.index(a.home_nb.select()) == 1
    assert a.home_title_label.cget("text") == "收件人库"
    assert a._nav_buttons["pool"].cget("style") == "NavOn.TButton"
    a.show_home_section("runs")
    a.root.update()
    assert a.home_title_label.cget("text") == "执行历史"
    a.show_screen("settings")
    a.root.update()
    assert a._nav_buttons["settings"].cget("style") == "NavOn.TButton"
    assert a.root.tk.call("ttk::style", "theme", "use") == "clam"
    labels = [str(w.cget("text")) for w in _walk(a.root)
              if _cls(w) in ("TLabel", "Label")]
    assert "工作台" in labels and "主导航" not in labels

# ---------------------------------------------------------------- v3.5.1

def test_settings_page_clean_and_default_outdir(app):
    """设置页：删除 GP 统一口令与全局文件名模板；默认输出目录显示程序目录。"""
    a, _mb, _fd, _err = app
    a.show_screen("settings")
    a.root.update()
    labels = [str(w.cget("text")) for w in _walk(a._screens["settings"])
              if _cls(w) == "TLabel"]
    joined = "\n".join(labels)
    assert "GP" not in joined
    assert "文件名模板：" not in joined
    assert not hasattr(a, "set_fname")
    assert not hasattr(a, "set_gp_mode")
    assert a.set_outdir.get() == str(lp_tool.app_dir())
    a.set_outdir.set(str(lp_tool.app_dir()))
    a._save_settings()
    st, _e = a.store.load_settings()
    assert st.get("outdir") == "", "程序目录应存为空值，便携包换位置仍默认跟随"


def test_filename_template_per_word_mailmerge_item(app):
    """v3.5.1：文件名模板是 Word-MailMerge 条目参数，不再在设置页。"""
    a, _mb, _fd, _err = app
    ev = _save_event(a, "文件名", [{"name": "甲", "to": "a@x.com",
                                    "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "generate", "params": {}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    labels = [str(w.cget("text")) for w in _walk(a._wb_params_frame) if _cls(w) == "TLabel"]
    assert any("生成文件名模板" in x for x in labels), labels
    item = a._wb_event["pipeline"][0]
    item["params"]["filename_tpl"] = "{{LP名称}}_专属账单"
    a._wb_save()
    # 条目级模板覆盖事件默认
    unit = {"rep": {"name": "甲"}, "recipients": [], "lp_name": "甲",
            "group_key": "甲", "group": ""}
    name = lp_v3.item_output_name(a._wb_event, unit, "C:/x/模板.docx",
                                  filename_tpl=item["params"]["filename_tpl"])
    assert name == "甲_专属账单"
    saved, _e = a.store.load_event(ev["id"])
    assert saved["pipeline"][0]["params"]["filename_tpl"] == "{{LP名称}}_专属账单"

# ---------------------------------------------------------------- v3.6.0

def test_workbench_four_steps_no_attachments_or_run_page(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "四步流程", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a.root.update()
    titles = [str(b.cget("text")).splitlines()[0] for b in a._wb_step_buttons]
    assert len(titles) == 4, titles
    assert "附件" not in titles[2] and "预览与生成" not in titles
    assert "邮件模板" in titles[2] and "总览" in titles[3]
    assert not hasattr(a, "_wb_step_attachments")
    assert not hasattr(a, "_wb_step_run")


def test_process_stays_on_pipeline_and_uses_default_outdir(app, tmp_path, monkeypatch):
    a, _mb, _fd, _err = app
    pdf = tmp_path / "附件.pdf"
    import fitz
    doc = fitz.open(); doc.new_page(); doc.save(str(pdf)); doc.close()
    outdir = tmp_path / "统一输出"
    a.settings["outdir"] = str(outdir)
    a.store.save_settings(a.settings)
    ev = _save_event(a, "处理日志", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": [str(pdf)]}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    a._wb_process()
    for _ in range(600):
        a.root.update()
        if not a.running:
            break
        import time as _t
        _t.sleep(0.02)
    for _ in range(20):
        a.root.update()
    assert a._wb_step == 1, "处理不应跳到总览/其它页面"
    logtxt = a._wb_pipe_log.get("1.0", "end")
    assert "[附件重复提示]" not in logtxt
    assert any(line.startswith("[20") for line in logtxt.splitlines() if line)
    saved, _e = a.store.load_event(ev["id"])
    assert saved["options"]["outdir"] == str(outdir)
    paths = [p for v in (saved.get("process_results") or {}).values() for p in v]
    assert paths and all(str(outdir) in p for p in paths), paths


def test_log_duplicate_notice_orange(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "日志颜色", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    a._wb_log("[附件重复提示] 文件 A 与 B 内容相同")
    a.root.update()
    box = a._wb_pipe_log
    assert box.tag_cget("dup", "foreground") == "#c25e00"
    assert box.tag_ranges("dup"), "重复提示应打上橙色 tag"


def test_mail_pregen_enters_overview(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "预生成", [{"name": "甲", "to": "a@x.com",
                                    "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a.root.update()
    a._wb_mail_preview_next()
    a.root.update()
    assert a._wb_step == 3
    assert hasattr(a, "_wb_ov_rows")


def test_overview_attachment_filename_opens(app, tmp_path, monkeypatch):
    a, _mb, _fd, _err = app
    pdf = tmp_path / "打开我.pdf"
    import fitz
    doc = fitz.open(); doc.new_page(); doc.save(str(pdf)); doc.close()
    ev = _save_event(a, "附件打开", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": [str(pdf)]}}])
    ev["process_results"] = {"甲": [str(pdf)]}
    a.store.save_event(ev)
    a.open_event(ev["id"])
    a._wb_show_step(3)
    a.root.update()
    opened = []
    monkeypatch.setattr(lp_tool.os, "startfile", lambda p: opened.append(p))
    labels = [w for w in _walk(a._wb_ov_files_frame)
              if _cls(w) == "Label" and str(w.cget("text")).startswith("· ")]
    assert labels, "总览附件栏应显示文件名"
    assert labels[0].bind("<Button-1>"), "附件文件名应绑定点击打开"
    a._wb_open_file(str(pdf))
    a.root.update()
    assert opened == [str(pdf)]


def test_top_actions_only_in_event(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "顶部按钮", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}])
    a.show_screen("home")
    a.root.update()
    assert a._ev_actions.winfo_manager() == ""
    a.open_event(ev["id"])
    a.root.update()
    assert a._ev_actions.winfo_manager() == "pack"
    texts = [str(w.cget("text")) for w in _walk(a._ev_actions) if _cls(w) == "TButton"]
    assert "返回列表" in texts and "事件信息…" in texts


def test_home_edit_event_dialog_renames(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "旧名字", [{"name": "甲", "to": "a@x.com",
                                    "open_pw": "Open123456", "owner_pw": "Owner123456"}])
    a.show_screen("home")
    a.root.update()
    ids = a.ev_tree.get_children()
    target = next(i for i in ids if a.ev_tree.item(i)["values"][0] == "旧名字")
    a.ev_tree.selection_set(target)
    a._edit_event_dialog()
    a.root.update()
    dlg = _find_toplevel(a.root, "事件信息")
    assert dlg is not None
    ents = [w for w in _walk(dlg) if _cls(w) == "TEntry"]
    ents[0].delete(0, "end")
    ents[0].insert(0, "首页新名字")
    for t, b in _buttons(dlg):
        if t == "保存":
            b.invoke()
            break
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["name"] == "首页新名字"


def test_event_info_send_mode_actually_effective(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "模式确认", [
        {"name": "甲", "to": "a@x.com", "group": "G1",
         "open_pw": "Open123456", "owner_pw": "Owner123456"},
        {"name": "乙", "to": "b@x.com", "group": "G1",
         "open_pw": "Open223456", "owner_pw": "Owner223456"},
    ], pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_event_info()
    a.root.update()
    dlg = _find_toplevel(a.root, "事件信息")
    combo = next(w for w in _walk(dlg) if _cls(w) == "TCombobox")
    combo.set("每人一份")
    combo.event_generate("<<ComboboxSelected>>")
    for t, b in _buttons(dlg):
        if t == "保存":
            b.invoke()
            break
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["options"]["send_mode"] == "per_lp"
    assert len(lp_v3.draft_units(saved)) == 2, "每人一份应产生 2 封"
    # 切回按组合并
    a._wb_event_info()
    a.root.update()
    dlg = _find_toplevel(a.root, "事件信息")
    combo = next(w for w in _walk(dlg) if _cls(w) == "TCombobox")
    combo.set("按组合并")
    combo.event_generate("<<ComboboxSelected>>")
    for t, b in _buttons(dlg):
        if t == "保存":
            b.invoke()
            break
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["options"]["send_mode"] == "per_group"
    assert len(lp_v3.draft_units(saved)) == 1, "同组合并应产生 1 封"

# ---------------------------------------------------------------- v3.7.0

def test_no_progress_bars_anymore(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "无进度条", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    assert not hasattr(a, "wb_progress")
    assert not hasattr(a, "_wb_progress2")
    a._wb_show_step(3)
    a.root.update()


def test_mail_editor_wysiwyg_toolbar_no_preview(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "编辑器", [{"name": "甲", "to": "a@x.com",
                                    "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a.root.update()
    assert not hasattr(a, "preview_label"), "右侧预览已删除"
    assert hasattr(a, "body_text")
    labels = [str(w.cget("text")) for w in _walk(a._wb_content)
              if _cls(w) in ("TLabel", "TButton", "TCombobox", "TMenubutton", "Button")]
    for need in ("字体", "字号", "文字色", "背景色", "左对齐", "居中",
                 "右对齐", "插入图片…", "插入字段"):
        assert any(need in x for x in labels), "缺少编辑器控件：%s" % need
    assert not any(str(w.cget("text")) == "设定" for w in _walk(a._wb_content)
                   if _cls(w) == "TButton"), "「设定」按钮应删除"


def test_editor_image_insert_and_html_roundtrip(app, tmp_path):
    a, _mb, _fd, _err = app
    from PIL import Image
    png = tmp_path / "pic.png"
    Image.new("RGB", (40, 30), (200, 100, 50)).save(str(png))
    ev = _save_event(a, "图片", [{"name": "甲", "to": "a@x.com",
                                  "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a.root.update()
    a.body_text.insert("end", "文字")
    lp_tool._editor_insert_image(a.body_text, str(png))
    a.root.update()
    html = lp_tool.text_to_html(a.body_text)
    assert "<img" in html and "data:image/png;base64," in html
    assert a.body_text.image_names(), "编辑器中应显示图片"
    lp_tool.load_html(a.body_text, html)
    a.root.update()
    assert a.body_text.image_names(), "从 HTML 重新加载后图片仍应显示"


def test_overview_editor_saves_per_email_override(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "单封编辑", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(3)
    a.root.update()
    assert hasattr(a, "_ov_body_text")
    key = a._wb_ov_cur_key
    a._ov_subject_var.set("个性化主题")
    a._ov_body_text.delete("1.0", "end")
    a._ov_body_text.insert("1.0", "个性化正文")
    a._wb_ov_save()
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    ov = (saved.get("mail_overrides") or {}).get(key)
    assert ov and ov["subject"] == "个性化主题"
    assert "个性化正文" in ov["body"]


def test_run_event_uses_per_email_override(tmp_path, monkeypatch):
    import fitz
    pdf = tmp_path / "a.pdf"
    doc = fitz.open(); doc.new_page(); doc.save(str(pdf)); doc.close()
    ev = lp_v3.new_event("覆盖生成")
    ev["recipients"] = [lp_v3.normalize_recipient(
        {"name": "甲", "to": "a@x.com", "open_pw": "Open123456", "owner_pw": "Owner123456"})]
    ev["pipeline"] = [lp_v3.normalize_item({"kind": "attach",
                                             "params": {"files": [str(pdf)]}})]
    key = next(iter(dict(lp_v3.draft_units(ev))))
    ev["mail_overrides"] = {key: {"subject": "单独主题", "body": "<p>单独正文</p>"}}
    calls = []
    monkeypatch.setattr(lp_v3, "create_outlook_draft",
                        lambda rec, files, subject, body, *a, **k: calls.append((subject, body)))
    store = lp_v3.Store(tmp_path / "data")
    run = lp_v3.run_event(store, ev, outdir=tmp_path / "out",
                          create_drafts=True, word=object())
    assert run is not None
    assert calls and calls[0][0] == "单独主题"
    assert "单独正文" in calls[0][1]

# ---------------------------------------------------------------- v3.7.2

def test_recipient_table_hides_fields_but_edit_preserves_them(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "隐藏字段", [{
        "name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
        "owner_pw": "Owner123456",
        "fields": {"Call_Amount": "200,000", "Due_Date": "2026-11-15"},
    }])
    a.open_event(ev["id"])
    a._wb_show_step(0)
    a.root.update()
    heads = [a._wb_rec_tree.heading(c)["text"] for c in a._wb_rec_tree["columns"]]
    assert "Call_Amount" not in heads and "Due_Date" not in heads
    # 打开编辑弹窗再保存，不应清空导入的字段
    tree = a._wb_rec_tree
    tree.selection_set(tree.get_children()[0])
    a._wb_edit_recipient()
    a.root.update()
    dlg = _find_toplevel(a.root, "编辑收件人")
    assert dlg is not None
    labels = [str(w.cget("text")) for w in _walk(dlg) if _cls(w) == "TLabel"]
    assert "Call_Amount" not in labels and "Due_Date" not in labels
    for t, b in _buttons(dlg):
        if t == "保存":
            b.invoke()
            break
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["recipients"][0]["fields"]["Call_Amount"] == "200,000"
    assert saved["recipients"][0]["fields"]["Due_Date"] == "2026-11-15"


def test_inserted_image_is_visible_in_editor(app, tmp_path):
    a, _mb, _fd, _err = app
    from PIL import Image
    png = tmp_path / "pic.png"
    Image.new("RGB", (80, 50), (200, 120, 90)).save(str(png))
    ev = _save_event(a, "图片可见", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a.root.update()
    a.body_text.insert("end", "前")
    lp_tool._editor_insert_image(a.body_text, str(png))
    a.root.update()
    names = a.body_text.image_names()
    assert names, "图片应插入编辑器"
    # 无头测试窗口里 bbox 可能拿不到；真实 GUI 的可见性由 e2e 的 bbox 校验覆盖。
    assert a.body_text.index(names[0]), "图片应插在有效文本索引"
    token = next((t for t in a.body_text.tag_names()
                  if t.startswith("__elide_")), None)
    assert token is not None, "图片 token 应被 elide 隐藏"
    assert str(a.body_text.cget("state")) == "normal"
    html = lp_tool.text_to_html(a.body_text)
    lp_tool.load_html(a.body_text, html)
    a.root.update()
    assert a.body_text.image_names(), "重新加载后图片仍可见"

# ---------------------------------------------------------------- v3.7.3

def _field_template_with_serial(app, tmp_path):
    import openpyxl
    tpl = tmp_path / "t3.docx"
    _mini_docx(tpl, "{{Due_Date}}")
    ev = _save_event(app, "日期选择", [
        {"name": "甲公司", "to": "a@x.com", "open_pw": "Open123456",
         "owner_pw": "Owner123456"}],
        pipeline=[{"kind": "generate", "params": {"template_path": str(tpl)}}])
    app.open_event(ev["id"])
    app._wb_show_step(1)
    app.root.update()
    src = tmp_path / "日期.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["收件人名称", "{{Due_Date}}"])
    ws.append(["甲公司", 46193])
    wb.save(str(src))
    return ev, app._wb_event["pipeline"][0], src


def test_import_field_template_keep_excel_display_text(app, tmp_path):
    a, _mb, fd, _err = app
    ev, item, src = _field_template_with_serial(a, tmp_path)
    fd.open_path = str(src)
    a._date_import_mode = lambda cols, count: "keep"
    a._wb_import_field_template(item)
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved["recipients"][0]["fields"]["Due_Date"] == "46193"


def test_import_field_template_date_format_options(app, tmp_path):
    from pathlib import Path as _P
    a, _mb, fd, _err = app
    # 直接验证格式化函数
    assert a._format_serial_date("46193", "iso") == "2026-06-20"
    assert a._format_serial_date("46193", "cn") == "2026年06月20日"
    assert a._format_serial_date("46193", "slash") == "2026/06/20"
    assert a._format_serial_date("46193", "keep") == "46193"
    assert a._format_serial_date("1222", "iso") == "1222"

# ---------------------------------------------------------------- v3.8.0

def test_mail_page_button_visible_and_toolbar_not_duplicated(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "邮件页", [{"name": "甲", "to": "a@x.com",
                                    "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a.root.update()
    btn_texts = [str(w.cget("text")) for w in _walk(a._wb_content)
                 if _cls(w) == "TButton"]
    assert "保存模板" in btn_texts
    assert "预生成 →" not in btn_texts
    assert "预生成" in a._wb_next_btn.cget("text")
    font_labels = [w for w in _walk(a._wb_content)
                   if _cls(w) in ("TLabel",) and str(w.cget("text")).strip() == "字体"]
    assert len(font_labels) == 1, "工具栏不应重复"


def test_unsaved_mail_prompts_save_discard_cancel(app):
    a, mb, _fd, _err = app
    ev = _save_event(a, "未保存", [{"name": "甲", "to": "a@x.com",
                                    "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a.mail_subject_var.set("未保存主题")
    assert a._mail_dirty
    # 取消：留在当前页
    mb.answer3 = None
    a._wb_next_step()
    assert a._wb_step == 2 and a._mail_dirty
    # 保存：写入事件并进入下一步
    mb.answer3 = True
    a._wb_next_step()
    assert a._wb_step == 3 and not a._mail_dirty
    saved, _e = a.store.load_event(ev["id"])
    assert saved["mail"]["subject"] == "未保存主题"


def test_unsaved_mail_discard_keeps_old(app):
    a, mb, _fd, _err = app
    ev = _save_event(a, "丢弃", [{"name": "甲", "to": "a@x.com",
                                  "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a._wb_show_step(2)
    old = a._wb_event["mail"]["subject"]
    a.mail_subject_var.set("不要的主题")
    mb.answer3 = False
    a._wb_prev_step()
    assert a._wb_step == 1
    saved, _e = a.store.load_event(ev["id"])
    assert saved["mail"]["subject"] == old


def test_pregen_clears_per_email_overrides(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "预生成清覆盖", [{"name": "甲", "to": "a@x.com",
                                          "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    ev["mail_overrides"] = {"§甲": {"subject": "手改", "body": "<p>手改</p>"}}
    a.store.save_event(ev)
    a.open_event(ev["id"])
    a._wb_show_step(2)
    a._wb_mail_preview_next()
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    assert saved.get("mail_overrides") == {}


def test_step_prev_next_buttons(app):
    a, _mb, _fd, _err = app
    ev = _save_event(a, "步骤按钮", [{"name": "甲", "to": "a@x.com",
                                      "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                     pipeline=[{"kind": "attach", "params": {"files": []}}])
    a.open_event(ev["id"])
    a.root.update()
    assert a._wb_step == 0
    assert "disabled" in a._wb_prev_btn.state()
    a._wb_next_step()
    a.root.update()
    assert a._wb_step == 1
    a._wb_prev_step()
    a.root.update()
    assert a._wb_step == 0
