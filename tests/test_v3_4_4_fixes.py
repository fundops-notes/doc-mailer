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

    def showinfo(self, title, msg="", **k):
        self.calls.append(("info", str(title), str(msg)))

    def showwarning(self, title, msg="", **k):
        self.calls.append(("warning", str(title), str(msg)))

    def showerror(self, title, msg="", **k):
        self.calls.append(("error", str(title), str(msg)))

    def askyesno(self, title, msg="", **k):
        self.calls.append(("ask", str(title), str(msg)))
        return True


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


def test_version_is_349():
    assert lp_tool.APP_VERSION == "3.4.9"
    assert lp_v3.APP_VERSION == "3.4.9"


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
    assert "附件路径" not in [tree.heading(c)["text"] for c in cols]
    assert vals[cols.index("Call_Amount")] == "200,000"
    assert vals[cols.index("Due_Date")] == "2026-11-15"
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
    assert _descendant(a._wb_process_btn, right), "处理按钮必须在右栏"
    assert _descendant(a._uni_btn_holder, right), "统一附件折叠按钮必须在右栏"
    assert a._wb_pipe_tree.winfo_manager() == "pack", "条目表必须真正挂在左栏里"
    # ⑥ 横向滚动条
    assert str(a._wb_pipe_tree.cget("xscrollcommand")), "流水线条目表必须有横向滚动条"
    horiz = [w for w in _walk(a._wb_pipe_tree.master)
             if _cls(w) == "TScrollbar" and str(w.cget("orient")) == "horizontal"]
    assert horiz, "流水线条目表容器里应有横向滚动条控件"
    # 统一附件默认收起、可展开收起
    assert a._uni_frame.winfo_manager() == ""
    _buttons(a._wb_content, "📎 统一附件")[0][1].invoke()
    a.root.update()
    assert a._uni_frame.winfo_manager() == "pack"
    _buttons(a._wb_content, "📎 统一附件")[0][1].invoke()
    a.root.update()
    assert a._uni_frame.winfo_manager() == ""


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
                     pipeline=[{"kind": "attach", "params": {"files": [str(att)]}}])
    a.open_event(ev["id"])
    a._wb_show_step(1)
    a.root.update()
    tree = a._wb_pipe_tree
    tree.selection_set(tree.get_children()[0])
    a.root.update()
    assert any("导出路径模板" in t for t, _b in _buttons(a._wb_content))
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
    settings, _e = a.store.load_settings()
    assert settings["gp_owner_mode"] is True
    assert settings["gp_owner_pw"] == "GPsecret123"
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
    assert "产出颗粒度" in joined
    for stale in ("产出粒度", "附件路径", "路径来源", "生成文档"):
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
    ws.append(["甲公司", 200000, "甲方科技", "2026-11-15"])
    ws.append(["乙公司", 300000, "乙方科技", "2026-12-01"])
    ws.append(["查无此人", 999, "X", "2027-01-01"])
    wb.save(str(src))
    fd.open_path = str(src)
    a._wb_import_field_template(item)
    a.root.update()
    saved, _e = a.store.load_event(ev["id"])
    r1 = saved["recipients"][0]
    r2 = saved["recipients"][1]
    assert r1["fields"]["Call_Amount"] == "200000"
    assert r1["fields"]["Client_Name"] == "甲方科技"
    assert r2["fields"]["Call_Amount"] == "300000"
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
    assert "收件人字段：Call_Amount" in src_box.get()
    # 切换到第三个占位符 → 来源下拉要联动
    ph_box.set("{{Due_Date}}")
    ph_box.event_generate("<<ComboboxSelected>>")
    a.root.update()
    assert "日期" in src_box.get(), "切换占位符后来源下拉应跟着刷新，实际=%r" % src_box.get()
    # 选一个来源并确定
    want = [v for v in src_box.cget("values") if "收件人字段：Call_Amount" in v][0]
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
        z.writestr("app/lp_tool.py", "print('new version')\n")
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
    assert "robocopy" in bat and "lp_data" in bat and "start" in bat
    assert not (appdir / ".update_pkg.zip").exists(), "下载包应在解压后删除"
