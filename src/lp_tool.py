#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LP 文档分发助手（GUI exe）

功能：
1) LP 信息管理：名称、水印名称、收件人(To+CC)、邮件标题、打开口令、权限口令；
   支持从 Excel(xlsx)/CSV 批量导入，可导出导入模板
2) 邮件模板：标题 + HTML 正文，支持 {{字段}} 占位符批量区分
   可用字段：{{LP名称}} {{水印名称}} {{打开口令}} {{日期}}
   出于安全设计，{{权限口令}} 不是可用字段——权限口令永远不会出现在邮件里
3) 批量处理：每个 LP 独立水印 + 独立打开口令 + 独立权限口令（AES-256 加密）
4) Outlook 草稿：逐 LP 生成草稿并插入附件，只 Save 不 Send，
   最终发送必须人工在 Outlook 里复核后手动发出

数据文件：与本程序同目录的 lp_data.json（含全部口令，属敏感文件，请妥善保管）
"""

import json
import random
import re
import string
import sys
import threading
import webbrowser
from datetime import date
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import fitz
import watermark_encrypt as wm

APP_TITLE = "LP 文档分发助手"
FIELDS = ["{{LP名称}}", "{{水印名称}}", "{{打开口令}}", "{{日期}}"]

DEFAULT_SUBJECT = "基金运营文件（{{LP名称}} 专属）"
DEFAULT_BODY = """<div style="font-family:微软雅黑,Segoe UI,Arial;font-size:14px;color:#333;line-height:1.8;">
<p><b>{{LP名称}}</b> 您好：</p>
<p>附件为贵司专属的基金文件（已加密并加盖专属水印），请使用以下打开口令查阅：</p>
<p style="padding:10px 16px;background:#f5f6f8;border-left:3px solid #4a6fa5;border-radius:4px;">
<b>打开口令：{{打开口令}}</b></p>
<p style="color:#888;font-size:12px;">本文件带有贵司专属水印标识，仅供贵司内部使用，请妥善保管、请勿转发。</p>
<p>如有任何疑问，请随时与我们联系。<br><br>此致<br>【基金管理人名称】<br>{{日期}}</p>
</div>"""


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# 演示数据（均为虚构，邮箱用保留域名，可直接导入测试全流程）
MOCK_LPS = [
    ["苏州启衡创业投资合伙企业（有限合伙）", "", "ir@qiheng-demo.cn", "fin@qiheng-demo.cn", "", "", ""],
    ["宁波梅山保税港区晨曦汐诺股权投资合伙企业（有限合伙）", "", "contact@chenxi-demo.cn", "", "", "Demo-Open-88", ""],
    ["新疆乌鲁木齐市泓远红杉股权投资合伙企业（有限合伙）", "", "admin@hongyuan-demo.cn", "ir@hongyuan-demo.cn", "", "", ""],
    ["上海临港科创一期私募投资基金合伙企业（有限合伙）", "", "lg@chuangxin-demo.cn", "", "临港科创一期2026年度文件", "", ""],
    ["珠海横琴粤澳深谋股权投资合伙企业（有限合伙）", "", "gp@henqin-demo.cn", "fin@henqin-demo.cn", "", "", "Demo-Owner-66"],
    ["北京中关军融创启股权投资合伙企业（有限合伙）", "", "zg@rongchuang-demo.cn", "", "", "", ""],
    ["深圳前海深港现代服务业合作区汇金聚富投资合伙企业（有限合伙）", "演示专用水印", "huijin@qianhai-demo.cn", "", "", "", ""],
    ["杭州西湖区云栖数字经济股权投资合伙企业（有限合伙）", "", "yunqi@xihu-demo.cn", "ir@xihu-demo.cn", "", "", ""],
]
MOCK_HEADERS = ["LP名称", "水印名称", "收件人", "抄送", "邮件标题", "打开口令", "权限口令"]


DATA_FILE = app_dir() / "lp_data.json"

# ---------- 纯逻辑函数（便于脱离 GUI 测试） ----------

PW_CHARS = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"


def gen_pw(n=10):
    return "".join(random.choice(PW_CHARS) for _ in range(n))


def sanitize_filename(s, maxlen=40):
    s = re.sub(r'[\\/:*?"<>|]', "_", s).strip()
    return s[:maxlen] if s else "unnamed"


def split_mails(s):
    for ch in ["；", ";", "，", ",", " "]:
        s = s.replace(ch, ";")
    return [x for x in s.split(";") if x]


HEADER_ALIASES = {
    "name": {"lp名称", "名称", "主体名称", "lp name", "name"},
    "wm": {"水印名称", "水印"},
    "to": {"收件人", "收件邮箱", "to", "邮箱", "email"},
    "cc": {"抄送", "cc"},
    "subject": {"邮件标题", "标题", "主题", "subject"},
    "open_pw": {"打开口令", "打开密码", "口令", "密码", "open pw"},
    "owner_pw": {"权限口令", "权限密码", "owner pw"},
}


def parse_import_rows(rows):
    """rows: 二维列表，第一行为表头。返回 (LP记录列表, 未识别的列名列表)。"""
    if not rows:
        return [], []
    header = ["" if h is None else str(h).strip().lower() for h in rows[0]]
    col_map = {}
    unknown = []
    for idx, h in enumerate(header):
        hit = None
        for key, aliases in HEADER_ALIASES.items():
            if h in aliases:
                hit = key
                break
        if hit:
            col_map[hit] = idx
        elif h:
            unknown.append(str(rows[0][idx]))
    lps = []
    for row in rows[1:]:
        def cell(key):
            i = col_map.get(key)
            if i is None or i >= len(row):
                return ""
            v = row[i]
            return "" if v is None else str(v).strip()
        if not cell("name"):
            continue
        lps.append({
            "name": cell("name"),
            "wm": cell("wm") or cell("name"),
            "to": cell("to"),
            "cc": cell("cc"),
            "subject": cell("subject"),
            "open_pw": cell("open_pw") or gen_pw(),
            "owner_pw": cell("owner_pw") or gen_pw(),
        })
    return lps, unknown


def read_spreadsheet(path):
    p = Path(path)
    if p.suffix.lower() == ".csv":
        import csv
        with open(p, newline="", encoding="utf-8-sig") as f:
            return [row for row in csv.reader(f)]
    else:
        import openpyxl
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.active
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
            if len(rows) > 5000:
                break
        wb.close()
        return rows


def _normalize_word_html(html):
    """Word 另存的 HTML 会把 {{字段}} 按语言标记拆进多个 span，
    导致占位符替换失效——这里把字段内部夹杂的标签剥掉，让字段复原成连续文本。"""
    def rejoin(m):
        return re.sub(r"<[^>]+>", "", m.group(0))
    return re.sub(r"\{\{[^}]*?\}\}", rejoin, html)


def render_template(tpl, rec):
    mapping = {
        "{{LP名称}}": rec.get("name", ""),
        "{{水印名称}}": rec.get("wm", ""),
        "{{打开口令}}": rec.get("open_pw", ""),
        "{{日期}}": date.today().strftime("%Y年%m月%d日"),
    }
    for k, v in mapping.items():
        tpl = tpl.replace(k, v)
    return tpl


def create_outlook_draft(rec, attachments, subject_tpl, body_tpl):
    """在 Outlook 草稿箱创建邮件（只 Save，绝不 Send）。"""
    import win32com.client as win32
    ol = win32.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)  # 0 = olMailItem
    mail.To = ";".join(split_mails(rec.get("to", "")))
    cc = split_mails(rec.get("cc", ""))
    if cc:
        mail.CC = ";".join(cc)
    mail.Subject = render_template(subject_tpl, rec)
    mail.HTMLBody = render_template(body_tpl, rec)
    for f in attachments:
        mail.Attachments.Add(str(Path(f).resolve()))  # Outlook 是独立进程，必须绝对路径
    mail.Save()  # 只存草稿；不调用 Send
    return mail


def _load_data_file(path):
    """读取数据 json；不存在或损坏返回 None。"""
    try:
        if Path(path).exists():
            d = json.loads(Path(path).read_text(encoding="utf-8"))
            if isinstance(d, dict):
                return d
    except Exception:
        pass
    return None


def render_preview_pixmap(html, width=620, dpi=110):
    """用 PyMuPDF Story 引擎把 HTML 渲染成 tk 可用的 PNG 字节（近似预览）。
    挂载系统微软雅黑以支持中文粗体；无该字体时退回内置字体。"""
    import os
    import tempfile
    css = ('@font-face {font-family:cn; src:url(msyh.ttc);}'
           '@font-face {font-family:cn; src:url(msyhbd.ttc); font-weight:bold;}'
           '@font-face {font-family:微软雅黑; src:url(msyh.ttc);}'
           '@font-face {font-family:微软雅黑; src:url(msyhbd.ttc); font-weight:bold;}'
           '@font-face {font-family:"Microsoft YaHei"; src:url(msyh.ttc);}'
           '@font-face {font-family:"Microsoft YaHei"; src:url(msyhbd.ttc); font-weight:bold;}'
           'body {font-family:cn;}')
    try:
        story = fitz.Story(html=html, user_css=css,
                           archive=fitz.Archive("C:/Windows/Fonts"))
    except Exception:
        story = fitz.Story(html=html)
    fd, path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    try:
        writer = fitz.DocumentWriter(path)
        pad = 12
        height = 2400  # 长模板先给足高度，按实际内容裁切
        mediabox = fitz.Rect(0, 0, width, height)
        dev = writer.begin_page(mediabox)
        more, filled = story.place(mediabox + (pad, pad, -pad, -pad))
        story.draw(dev)
        writer.end_page()
        writer.close()
        doc = fitz.open(path)
        page = doc[0]
        # 找到内容实际底边，裁掉大片空白
        blocks = [b for b in page.get_text("blocks") if b[4].strip()]
        bottom = max((b[3] for b in blocks), default=height // 2) + pad * 2
        clip = fitz.Rect(0, 0, width, min(bottom + 8, height))
        pix = page.get_pixmap(dpi=dpi, clip=clip)
        ret = pix.tobytes("png")
        doc.close()
        return ret
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


# ---------- GUI ----------

class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("980x680")
        try:
            root.state("zoomed")
        except Exception:
            pass

        self.data = self.load_data()
        self.pdfs = []
        self.running = False
        self._after_id = None
        self._preview_img = None

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb
        self.tab_lp = ttk.Frame(nb)
        self.tab_tpl = ttk.Frame(nb)
        self.tab_run = ttk.Frame(nb)
        nb.add(self.tab_lp, text=" ① LP 信息管理 ")
        nb.add(self.tab_tpl, text=" ② 邮件模板 ")
        nb.add(self.tab_run, text=" ③ 生成文件 / Outlook 草稿 ")

        self.build_tab_lp()
        self.build_tab_tpl()
        self.build_tab_run()
        nb.bind("<<NotebookTabChanged>>",
                lambda e: self.refresh_preview()
                if nb.index(nb.select()) == 1 and hasattr(self, "preview_label") else None)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ----- 数据 -----

    def load_data(self):
        d = _load_data_file(DATA_FILE)
        if d is None:
            # 首次运行：若旁边有示例数据（lp_data.example.json），载入演示 LP
            example = DATA_FILE.parent / "lp_data.example.json"
            d = _load_data_file(example) or {
                "lps": [], "subject": DEFAULT_SUBJECT, "body": DEFAULT_BODY}
            try:  # 直接落盘为 lp_data.json，示例文件保持只读不被改动
                DATA_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=2),
                                     encoding="utf-8")
            except Exception:
                pass
        return d

    def save_data(self):
        try:
            DATA_FILE.write_text(json.dumps(self.data, ensure_ascii=False, indent=2),
                                 encoding="utf-8")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法写入 {DATA_FILE}\n{e}")

    # ----- Tab1: LP 管理 -----

    def build_tab_lp(self):
        f = self.tab_lp
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=4, padx=6)
        for text, cmd in [
            ("导入 Excel/CSV", self.do_import),
            ("导出空白模板", self.export_template),
            ("导出示例清单", self.export_sample),
            ("新增", lambda: self.edit_dialog(None)),
            ("编辑", lambda: self.edit_selected()),
            ("删除", self.delete_selected),
            ("补齐缺失口令", self.fill_passwords),
        ]:
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=3)

        cols = ["name", "wm", "to", "cc", "open_pw", "owner_pw", "subject"]
        heads = ["LP名称", "水印名称", "收件人(To)", "抄送(CC)",
                 "打开口令", "权限口令(勿发)", "邮件标题(空=用模板)"]
        self.tree = ttk.Treeview(f, columns=cols, show="headings")
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=150 if c in ("to", "subject") else 110)
        self.tree.pack(fill="both", expand=True, padx=6, pady=4)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

        ttk.Label(f, foreground="#b00",
                  text="安全提示：本程序同目录的 lp_data.json 含全部权限口令，属敏感文件，请限定在本机/受控共享目录使用。"
                  ).pack(fill="x", padx=8, pady=4)
        self.refresh_tree()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for rec in self.data["lps"]:
            self.tree.insert("", "end", values=[rec.get(k, "") for k in
                                                ["name", "wm", "to", "cc",
                                                 "open_pw", "owner_pw", "subject"]])

    def selected_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree.index(sel[0])

    def do_import(self):
        path = filedialog.askopenfilename(title="选择 LP 清单",
                                          filetypes=[("表格", "*.xlsx *.csv"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            rows = read_spreadsheet(path)
            lps, unknown = parse_import_rows(rows)
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        if not lps:
            messagebox.showwarning("导入", "没有解析到任何 LP 记录（需要表头 + 名称列）")
            return
        names = {r["name"] for r in self.data["lps"]}
        added = 0
        for rec in lps:  # 同名更新，异名新增
            if rec["name"] in names:
                self.data["lps"] = [rec if r["name"] == rec["name"] else r
                                    for r in self.data["lps"]]
            else:
                self.data["lps"].append(rec)
                added += 1
        self.save_data()
        self.refresh_tree()
        msg = f"导入 {len(lps)} 条（新增 {added}，其余覆盖同名记录）。"
        if unknown:
            msg += f"\n未识别的列（已忽略）：{'、'.join(unknown)}"
        messagebox.showinfo("导入完成", msg)

    def export_template(self):
        self._write_xlsx("LP清单导入模板.xlsx", [MOCK_HEADERS,
                       ["示例：苏州某某投资合伙企业（有限合伙）", "（留空=同LP名称）",
                        "a@x.com;b@x.com", "cc@x.com", "（留空=用邮件模板）",
                        "（留空=自动生成）", "（留空=自动生成，勿发给LP）"]],
                       "模板已导出", "表头列名即导入识别依据，可增删其他列。\n口令列留空会自动生成随机口令。")

    def export_sample(self):
        self._write_xlsx("LP清单示例（演示数据）.xlsx", [MOCK_HEADERS] + MOCK_LPS,
                         "示例清单已导出",
                         "共 8 条虚构演示数据，可直接「导入 Excel/CSV」体验全流程。\n"
                         "邮箱均为保留域名，不会发往真实地址（草稿也需人工确认才会发送）。")

    def _write_xlsx(self, default_name, rows, ok_title, ok_msg):
        path = filedialog.asksaveasfilename(title="保存", defaultextension=".xlsx",
                                            initialfile=default_name,
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        for row in rows:
            ws.append(row)
        for col, w in zip("ABCDEFG", [36, 22, 30, 18, 24, 16, 28]):
            ws.column_dimensions[col].width = w
        try:
            wb.save(path)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return
        messagebox.showinfo(ok_title, ok_msg)

    def edit_dialog(self, idx):
        rec = dict(self.data["lps"][idx]) if idx is not None else {
            "name": "", "wm": "", "to": "", "cc": "", "subject": "",
            "open_pw": gen_pw(), "owner_pw": gen_pw()}
        win = tk.Toplevel(self.root)
        win.title("编辑 LP" if idx is not None else "新增 LP")
        win.grab_set()
        entries = {}
        rows = [("LP名称", "name"), ("水印名称", "wm"),
                ("收件人To（多人用;分隔）", "to"), ("抄送CC（可空）", "cc"),
                ("邮件标题（空=用模板）", "subject"),
                ("打开口令（发LP）", "open_pw"), ("权限口令（勿发LP）", "owner_pw")]
        for i, (label, key) in enumerate(rows):
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="e", padx=6, pady=4)
            e = ttk.Entry(win, width=52, show="*" if key == "owner_pw" else "")
            e.insert(0, rec.get(key, ""))
            e.grid(row=i, column=1, padx=6, pady=4)
            entries[key] = e
        ttk.Label(win, text="权限口令输入框默认遮蔽，仅内部留档", foreground="#888").grid(
            row=len(rows), column=1, sticky="w")

        def ok():
            out = {k: e.get().strip() for k, e in entries.items()}
            if not out["name"]:
                messagebox.showwarning("必填", "LP名称不能为空", parent=win)
                return
            if out["owner_pw"] and out["owner_pw"] == out["open_pw"]:
                messagebox.showwarning("口令", "权限口令不能与打开口令相同", parent=win)
                return
            out["wm"] = out["wm"] or out["name"]
            if idx is None:
                self.data["lps"].append(out)
            else:
                self.data["lps"][idx] = out
            self.save_data()
            self.refresh_tree()
            win.destroy()

        btns = ttk.Frame(win)
        btns.grid(row=len(rows) + 1, column=0, columnspan=2, pady=8)
        ttk.Button(btns, text="保存", command=ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=8)

    def edit_selected(self):
        idx = self.selected_index()
        if idx is None:
            messagebox.showinfo("提示", "请先选中一行")
            return
        self.edit_dialog(idx)

    def delete_selected(self):
        idx = self.selected_index()
        if idx is None:
            return
        name = self.data["lps"][idx]["name"]
        if messagebox.askyesno("删除", f"确认删除「{name}」？"):
            del self.data["lps"][idx]
            self.save_data()
            self.refresh_tree()

    def fill_passwords(self):
        n = 0
        for rec in self.data["lps"]:
            if not rec.get("open_pw"):
                rec["open_pw"] = gen_pw(); n += 1
            if not rec.get("owner_pw"):
                rec["owner_pw"] = gen_pw(); n += 1
        self.save_data()
        self.refresh_tree()
        messagebox.showinfo("补齐口令", f"新生成 {n} 个口令（导入时留空的会已自动生成，此按钮用于兜底）")

    # ----- Tab2: 邮件模板 -----

    _word_app = None

    def build_tab_tpl(self):
        f = self.tab_tpl
        ttk.Label(f, text="邮件标题模板：").pack(anchor="w", padx=8, pady=(8, 0))
        self.subject_var = tk.StringVar(value=self.data.get("subject", DEFAULT_SUBJECT))
        ttk.Entry(f, textvariable=self.subject_var, width=110).pack(fill="x", padx=8)

        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=8, pady=6)
        ttk.Button(bar, text="在 Word 中编辑正文（所见即所得）",
                   command=self.edit_body_in_word).pack(side="left")
        ttk.Button(bar, text="刷新预览", command=self.refresh_preview).pack(side="left", padx=6)
        ttk.Button(bar, text="保存模板", command=self.save_template).pack(side="right")
        ttk.Button(bar, text="恢复默认", command=self.reset_template).pack(side="right", padx=2)

        ttk.Label(f, foreground="#666", justify="left", text=(
            "点上方按钮用 Word 编辑正文：可直接粘贴带格式的富文本（Word / 网页 / Excel 表格均可），"
            "用 Word 工具栏加粗、调色、插表格图片。\n"
            "可用字段（在正文里原样输入这些占位符）：{{LP名称}}  {{水印名称}}  {{打开口令}}  {{日期}}\n"
            "没有权限口令字段——权限口令按规则永远不进入邮件。"
        )).pack(anchor="w", padx=8, pady=(0, 4))

        self.preview_label = tk.Label(f, anchor="nw", background="#ffffff")
        self.preview_label.pack(fill="both", expand=True, padx=8, pady=4)

    def edit_body_in_word(self):
        import win32com.client as win32
        html_path = app_dir() / "邮件正文模板_编辑中.html"
        full = ('<!DOCTYPE html><html><head><meta charset="utf-8">'
                '<title>邮件正文模板</title></head><body>'
                + self.data.get("body", DEFAULT_BODY) + "</body></html>")
        html_path.write_text(full, encoding="utf-8")
        try:
            word = win32.DispatchEx("Word.Application")
            word.Visible = True
            doc = word.Documents.Open(str(html_path))
        except Exception as e:
            messagebox.showerror("需要 Word",
                                 f"无法用 Word 打开正文编辑：{e}\n请确认本机已安装 Word。")
            return
        self._word_app = (word, str(html_path))
        try:
            word.Activate()
        except Exception:
            pass
        win = tk.Toplevel(self.root)
        win.title("Word 编辑中")
        tk.Label(win, justify="left", padx=14, pady=10, text=(
            "已用 Word 打开模板正文。\n\n"
            "· 可直接粘贴富文本，用 Word 工具栏排版\n"
            "· 编辑完成后点【读取并更新模板】\n"
            "· 之后可关闭 Word")).pack()
        bb = ttk.Frame(win)
        bb.pack(pady=8)

        def do_read():
            if self._read_word_body():
                win.destroy()

        ttk.Button(bb, text="读取并更新模板", command=do_read).pack(side="left", padx=6)
        ttk.Button(bb, text="取消", command=win.destroy).pack(side="left", padx=6)

    def _read_word_body(self):
        import os
        import re
        import tempfile
        candidates = []
        if getattr(self, "_word_app", None):
            word, path = self._word_app
            try:
                for d in word.Documents:
                    if d.FullName.lower() == path.lower():
                        try:
                            d.WebOptions.Encoding = 65001  # UTF-8
                        except Exception:
                            pass
                        fd, tmp = tempfile.mkstemp(suffix=".html")
                        os.close(fd)
                        d.SaveAs2(tmp, 10)  # 10 = wdFormatFilteredHTML
                        candidates.append(Path(tmp))
                        import shutil
                        shutil.copyfile(tmp, path)  # 让原文档路径保持最新内容
                        break
            except Exception:
                pass  # Word 已被用户关闭等情况，走文件回退
        f = app_dir() / "邮件正文模板_编辑中.html"
        if f.exists():
            candidates.append(f)
        import re as _re
        for c in candidates:
            try:
                txt = c.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                txt = c.read_text(encoding="gbk", errors="ignore")
            m = _re.search(r"<body[^>]*>(.*)</body>", txt, _re.S | _re.I)
            if m and len(m.group(1).strip()) > 10:
                body = _normalize_word_html(m.group(1).strip())
                self.data["body"] = body
                self.save_data()
                self.refresh_preview()
                return True
        messagebox.showwarning(
            "未读取到内容",
            "没读到已保存的正文。请在 Word 中编辑后点【读取并更新模板】\n"
            "（若已关闭 Word，请先在 Word 里 Ctrl+S 保存过再试）。")
        return False

    def collect_template(self):
        return (self.subject_var.get().strip(),
                self.data.get("body", DEFAULT_BODY))

    def save_template(self):
        self.data["subject"], self.data["body"] = self.collect_template()
        self.save_data()
        self.refresh_preview()
        messagebox.showinfo("模板", "已保存")

    def reset_template(self):
        self.subject_var.set(DEFAULT_SUBJECT)
        self.data["body"] = DEFAULT_BODY
        self.save_data()
        self.refresh_preview()

    def refresh_preview(self):
        try:
            import io
            from PIL import Image, ImageTk
            idx = self.selected_index()
            rec = self.data["lps"][idx] if idx is not None and self.data["lps"] else {
                "name": "【示例LP名称】", "wm": "【示例水印名称】", "open_pw": "【示例口令】"}
            subject, body = self.collect_template()
            html = (f'<div style="font-family:sans-serif;">'
                    f'<p><b>标题：</b>{render_template(subject, rec)}</p><hr>'
                    f'{render_template(body, rec)}</div>')
            png = render_preview_pixmap(html)
            img = Image.open(io.BytesIO(png))
            w = max(self.preview_label.winfo_width(), 200) - 4
            if w > 20 and img.width > w:
                h = int(img.height * w / img.width)
                img = img.resize((w, h))
            self._preview_img = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self._preview_img, text="")
            self.preview_label.image = self._preview_img
        except Exception as e:
            self.preview_label.configure(
                image="", text=f"预览渲染失败：{e}\n（可点「刷新预览」重试）")

    def preview(self):
        idx = self.selected_index()
        rec = self.data["lps"][idx] if idx is not None and self.data["lps"] else {
            "name": "【示例LP名称】", "wm": "【示例水印名称】",
            "open_pw": "【示例口令】"}
        subject, body = self.collect_template()
        html = (f"<html><body><p><b>标题：</b>{render_template(subject, rec)}"
                f"</p><hr>{render_template(body, rec)}</body></html>")
        import tempfile
        p = Path(tempfile.gettempdir()) / "lp_mail_preview.html"
        p.write_text(html, encoding="utf-8")
        webbrowser.open(p.as_uri())

    # ----- Tab3: 生成 -----

    def build_tab_run(self):
        f = self.tab_run
        top = ttk.Frame(f)
        top.pack(fill="x", padx=8, pady=6)
        ttk.Button(top, text="添加原件PDF", command=self.add_pdf).pack(side="left")
        ttk.Button(top, text="移除选中", command=self.remove_pdf).pack(side="left", padx=4)
        ttk.Label(top, text="    输出目录：").pack(side="left")
        self.outdir_var = tk.StringVar(value=str(app_dir() / "out_wm"))
        ttk.Entry(top, textvariable=self.outdir_var, width=50).pack(side="left", padx=4)
        ttk.Button(top, text="浏览", command=self.browse_outdir).pack(side="left")

        self.pdf_list = tk.Listbox(f, height=5)
        self.pdf_list.pack(fill="x", padx=8, pady=4)

        opt = ttk.Frame(f)
        opt.pack(fill="x", padx=8, pady=2)
        ttk.Button(opt, text="水印设置…", command=self.open_wm_settings).pack(side="left")
        self.wm_summary = ttk.Label(opt, text="")
        self.wm_summary.pack(side="left", padx=10)
        ttk.Label(opt, foreground="#888",
                  text="水印文字 = 每个 LP 的「水印名称」列").pack(side="left")

        runbar = ttk.Frame(f)
        runbar.pack(fill="x", padx=8, pady=6)
        self.btn_files = ttk.Button(runbar, text="仅生成加密文件",
                                    command=lambda: self.run_batch(False))
        self.btn_files.pack(side="left")
        self.btn_draft = ttk.Button(runbar, text="生成加密文件 + Outlook草稿",
                                    command=lambda: self.run_batch(True))
        self.btn_draft.pack(side="left", padx=8)
        ttk.Label(runbar, foreground="#b00",
                  text="草稿只写入 Outlook 草稿箱，不会发送；发送前请人工复核。").pack(side="left")

        self.log_text = tk.Text(f, height=14, state="disabled",
                                font=("Consolas", 10), background="#fafafa")
        self.log_text.pack(fill="both", expand=True, padx=8, pady=4)
        self.update_wm_summary()

    def add_pdf(self):
        files = filedialog.askopenfilenames(title="选择原件 PDF",
                                            filetypes=[("PDF", "*.pdf")])
        for fp in files:
            if fp not in self.pdfs:
                self.pdfs.append(fp)
                self.pdf_list.insert("end", fp)

    def remove_pdf(self):
        for i in reversed(self.pdf_list.curselection()):
            del self.pdfs[i]
            self.pdf_list.delete(i)

    def browse_outdir(self):
        d = filedialog.askdirectory(title="输出目录")
        if d:
            self.outdir_var.set(d)

    # ----- 水印设置 -----

    LAYOUT_OPTS = [("自动（推荐）", "auto"), ("强制单行", "single"), ("强制折行", "wrap")]
    REPEAT_OPTS = [("单个", 1), ("三份（防裁剪）", 3), ("平铺 2列×3行", "tile")]
    ANGLE_OPTS = [("自动（沿对角线）", -1.0), ("45°", 45.0), ("30°", 30.0), ("水平", 0.0)]
    FONT_OPTS = [("默认", ""), ("黑体", "黑体"), ("楷体", "楷体"),
                 ("仿宋", "仿宋"), ("宋体", "宋体"), ("微软雅黑", "微软雅黑")]

    def update_wm_summary(self):
        c = dict(wm.WM_DEFAULTS, **(self.data.get("wm_settings") or {}))
        layout = dict(self.LAYOUT_OPTS).get(c.get("layout"), c.get("layout"))
        rep = {1: "单个", 3: "三份", "tile": "平铺"}.get(c.get("repeat"), c.get("repeat"))
        ang = "对角线" if float(c.get("angle", -1) or -1) < 0 else f"{c.get('angle'):.0f}°"
        fs = c.get("fs_max") or "自动"
        font = c.get("font") or "默认"
        self.wm_summary.configure(
            text=f"{layout} · {font} · 字号上限{fs} · {ang} · {rep} · "
                 f"浓度{float(c.get('opacity', 0.25)):.2f}")

    def open_wm_settings(self):
        cur = dict(wm.WM_DEFAULTS, **(self.data.get("wm_settings") or {}))
        win = tk.Toplevel(self.root)
        win.title("水印设置")
        win.grab_set()

        frm = ttk.Frame(win)
        frm.pack(padx=10, pady=8, fill="x")
        self._wm_vars = {}
        _row = [0]

        def add_row(label, widget, hint=""):
            r = _row[0]
            _row[0] += 1
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="e", padx=6, pady=3)
            widget.grid(row=r, column=1, sticky="w", padx=6, pady=3)
            if hint:
                ttk.Label(frm, foreground="#888", text=hint).grid(
                    row=r, column=2, sticky="w")

        # 布局
        lay = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.LAYOUT_OPTS])
        lay.set(dict((v, n) for n, v in self.LAYOUT_OPTS).get(cur.get("layout"), "自动（推荐）"))
        add_row("布局方式：", lay, hint="自动=可读则单行缩号，否则折行")
        self._wm_vars["layout"] = lay

        # 字体 / 字号 / 角度
        fnt = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.FONT_OPTS])
        fnt.set(dict((v, n) for n, v in self.FONT_OPTS).get(cur.get("font"), "默认"))
        add_row("字体：", fnt); self._wm_vars["font"] = fnt

        fsbox = ttk.Frame(frm)
        auto_fs = tk.BooleanVar(value=not cur.get("fs_max"))
        fs_spin_var = tk.IntVar(value=cur.get("fs_max") or 34)
        cbs = ttk.Checkbutton(fsbox, text="自动", variable=auto_fs)
        cbs.pack(side="left")
        sp = ttk.Spinbox(fsbox, from_=16, to=48, width=4, textvariable=fs_spin_var)
        sp.pack(side="left", padx=4)
        add_row("最大字号：", fsbox)
        self._wm_vars["_auto_fs"], self._wm_vars["fs_max"] = auto_fs, fs_spin_var

        ang = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.ANGLE_OPTS])
        cur_ang = float(cur.get("angle", -1) or -1)
        ang.set(dict((v, n) for n, v in self.ANGLE_OPTS).get(cur_ang, "自动（沿对角线）"))
        add_row("角度：", ang); self._wm_vars["angle"] = ang

        rep = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.REPEAT_OPTS])
        rep.set({1: "单个", 3: "三份（防裁剪）", "tile": "平铺 2列×3行"}
                .get(cur.get("repeat"), "单个"))
        add_row("份数：", rep); self._wm_vars["repeat"] = rep

        # 颜色 / 不透明度
        cbox = ttk.Frame(frm)
        color_var = tk.StringVar(value=cur.get("color") or "#737373")

        def pick_color():
            from tkinter import colorchooser
            c = colorchooser.askcolor(parent=win, title="水印颜色",
                                      initialcolor=color_var.get())
            if c and c[1]:
                color_var.set(c[1])

        ttk.Entry(cbox, textvariable=color_var, width=10).pack(side="left")
        ttk.Button(cbox, text="选色…", command=pick_color, width=6).pack(side="left", padx=4)
        add_row("颜色：", cbox); self._wm_vars["color"] = color_var

        opv = tk.DoubleVar(value=float(cur.get("opacity", 0.25)))
        op = ttk.Spinbox(frm, from_=0.05, to=0.50, increment=0.05, width=5, textvariable=opv)
        add_row("不透明度：", op); self._wm_vars["opacity"] = opv

        def collect():
            return {
                "layout": dict(self.LAYOUT_OPTS).get(lay.get(), "auto"),
                "font": dict(self.FONT_OPTS).get(fnt.get(), ""),
                "fs_max": 0 if auto_fs.get() else max(10, min(60, fs_spin_var.get())),
                "angle": dict(self.ANGLE_OPTS).get(ang.get(), -1.0),
                "repeat": dict(self.REPEAT_OPTS).get(rep.get(), 1),
                "color": color_var.get(),
                "opacity": float(opv.get()),
            }

        # 预览区
        ttk.Separator(win).pack(fill="x", padx=10)
        pf = ttk.Frame(win)
        pf.pack(padx=10, pady=6, fill="both", expand=True)
        self._wm_preview_label = tk.Label(pf, anchor="nw", background="#ffffff")
        self._wm_preview_label.pack(side="left", fill="both", expand=True)

        def refresh_preview():
            idx = self.selected_index()
            text = (self.data["lps"][idx].get("wm") if idx is not None
                    and self.data["lps"] else "新疆乌鲁木齐市泓远红杉股权投资合伙企业（有限合伙）")
            try:
                import io
                from PIL import Image, ImageTk
                doc = fitz.open()
                page = doc.new_page()  # A4
                wm.add_watermark(page, text, collect())
                png = page.get_pixmap(dpi=72).tobytes("png")
                img = Image.open(io.BytesIO(png))
                w = 340
                img = img.resize((w, int(img.height * w / img.width)))
                self._wm_preview_img = ImageTk.PhotoImage(img)
                self._wm_preview_label.configure(image=self._wm_preview_img, text="")
            except Exception as e:
                self._wm_preview_label.configure(image="", text=f"预览失败：{e}")

        ttk.Button(pf, text="刷新预览", command=refresh_preview).pack(side="left",
                                                                      padx=8, anchor="n")

        btns = ttk.Frame(win)
        btns.pack(pady=8)
        ttk.Button(btns, text="保存", command=lambda: save_and_close()).pack(
            side="left", padx=8)

        def save_and_close():
            self.data["wm_settings"] = collect()
            self.save_data()
            self.update_wm_summary()
            win.destroy()

        refresh_preview()

    def log(self, msg):
        def append():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, append)

    def run_batch(self, make_drafts):
        if self.running:
            return
        if not self.data["lps"]:
            messagebox.showwarning("提示", "LP 清单为空，请先在①中导入")
            return
        if not self.pdfs:
            messagebox.showwarning("提示", "请先添加原件 PDF")
            return
        if make_drafts:
            missing = [r["name"] for r in self.data["lps"] if not r.get("to")]
            if missing:
                messagebox.showwarning("缺少收件人",
                                       "以下 LP 没有收件人，将只生成文件、跳过草稿：\n" + "\n".join(missing))
        subject, body = self.collect_template()
        self.data["subject"], self.data["body"] = subject, body
        self.save_data()

        pdfs = list(self.pdfs)
        outdir = Path(self.outdir_var.get())
        outdir.mkdir(parents=True, exist_ok=True)
        cfg = dict(self.data.get("wm_settings") or {})
        lps = [dict(r) for r in self.data["lps"]]
        self.running = True
        self.btn_files.state(["disabled"])
        self.btn_draft.state(["disabled"])
        self.log(f"===== 开始：{len(lps)} 个 LP × {len(pdfs)} 个文件，"
                 f"{'含' if make_drafts else '不含'} Outlook 草稿 =====")

        def worker():
            try:
                for rec in lps:
                    outs = []
                    for src in pdfs:
                        out = outdir / f"{Path(src).stem}_{sanitize_filename(rec['name'])}.pdf"
                        try:
                            wm.process_one(Path(src), out, rec["wm"], rec["open_pw"],
                                           rec["owner_pw"], False, wm_cfg=cfg)
                            outs.append(out)
                            self.log(f"[文件OK] {out.name}")
                        except Exception as e:
                            self.log(f"[文件失败] {rec['name']} <- {Path(src).name}: {e}")
                    if make_drafts and outs and rec.get("to"):
                        try:
                            create_outlook_draft(rec, outs, subject, body)
                            self.log(f"[草稿OK] {rec['name']} -> Outlook 草稿箱（未发送）")
                        except Exception as e:
                            self.log(f"[草稿失败] {rec['name']}: {e}")
                    elif make_drafts and not rec.get("to"):
                        self.log(f"[跳过草稿] {rec['name']}（无收件人）")
                self.log("===== 全部完成 =====")
            finally:
                def done():
                    self.running = False
                    self.btn_files.state(["!disabled"])
                    self.btn_draft.state(["!disabled"])
                self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def on_close(self):
        self.data["subject"], self.data["body"] = self.collect_template()
        self.save_data()
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
