#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档分发助手 v3.4.4（GUI exe）

v3.4.4 修复（相对 v3.4.3）：
  · 流水线页左右分栏真正落盘：左栏=条目表+参数区，右栏=统一附件+处理按钮+对应表
  · 收件人表列值对齐（修复权限口令遮罩失效/字段错位）；编辑收件人保留旧附件数据
  · 收件人清单导入模板去掉「附件路径」列；导入该列只收进自定义字段
  · Word-MailMerge 参数区补「导出要素模板」「下载示例 Word 模板」；附加文件补「导出路径模板」
  · 「发送人统一密码」修复并接入引擎；流水线条目表补横向滚动条

v3.1 相对 v3.0 的结构变化：
  · 流水线从「六种步骤」重构为「两种产出条目」：
      ① Word-MailMerge（Word 模板 → 可选转 PDF / 加水印 / 加密）
      ② 附加文件（磁盘上已有的文件 → 可选加水印 / 加密）
    处理顺序固定（转 PDF → 加水印 → 加密），用户不用排；每条产出几份文件由条目自己的
    「产出颗粒度」决定。
  · 「数据来源」概念消解：来源下沉到字段级 —— 每个占位符单独选来源，并给预览值。
  · 工作台新增第 5 步「总览」：生成前逐人核对最终附件包与邮件样子。

v3.0 相对 v2.x 的结构变化：
  · 顶层从「标签页」改为「左侧固定导航 + 右侧内容区」
  · 引入 场景 Scenario（怎么做）/ 发送事件 Event（这次做给谁）两层结构
  · 数据从单文件 lp_data.json 改为目录 lp_data/（settings / events / scenarios / runs）

安全设计（不可动摇）：
  · 权限口令绝不进入邮件正文/标题/任何草稿字段
  · 打开口令 ≠ 权限口令
  · 只建草稿绝不发送（全库唯一 Outlook 写操作是 mail.Save()）
  · AES-256 加密，权限位仅打印
  · 草稿只写入 Outlook 草稿箱，发送前必须人工复核

数据文件：程序同目录的 lp_data/（含全部口令，属敏感文件，请妥善保管）
"""

import html as _htmlmod
import json
import os
import queue as _queue
import re
import sys
import threading
import webbrowser
from datetime import date, datetime
from html.parser import HTMLParser as _HTMLParser
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import fitz
import watermark_encrypt as wm

import lp_v3
from lp_v3 import (DEFAULT_FILENAME_TPL, GRANULARITIES, GRANULARITY_LABELS,
                   HEADER_ALIASES, ITEM_KIND_DESC, ITEM_KIND_LABELS, ITEM_KINDS,
                   MOCK_HEADERS, MOCK_LPS, PERMISSIONS, PERMISSION_LABELS,
                   Store, clone_event, gen_pw, migrate_v2_to_v3, needs_migration,
                   new_event, new_item, new_scenario, normalize_attachment,
                   normalize_event, normalize_item, normalize_recipient,
                   parse_import_rows, parse_import_rows_v3, pw_fingerprint,
                   read_spreadsheet, sanitize_filename, split_mails, split_paths,
                   sync_scenario_to_event)

APP_VERSION = "3.4.8"
# lp_v3.py 里也有一份版本号（导出备份会带上）。两处必须一致。
# 注意：发布检查会断言 lp_tool.py 里存在 'APP_VERSION = "3.4.8"' 字面量，别改成 import。
assert lp_v3.APP_VERSION == APP_VERSION, "版本号不一致：lp_tool.py 与 lp_v3.py 必须同步"
UPDATE_REPO = "fundops-notes/doc-mailer"  # GitHub 仓库（自动更新源）
APP_TITLE = f"文档分发助手 v{APP_VERSION}"

# 邮件模板里永远可用的内置字段（**没有权限口令**，这是架构保证不是靠自觉）
FIELDS = ["{{收件人名称}}", "{{水印名称}}", "{{打开口令}}", "{{日期}}"]

DEFAULT_SUBJECT = "机密文件（{{收件人名称}} 专属）"
DEFAULT_BODY = """<div style="font-family:微软雅黑,Segoe UI,Arial;font-size:14px;color:#333;line-height:1.8;">
<p><b>{{收件人名称}}</b> 您好：</p>
<p>附件为贵司专属的加密文件（已加盖专属水印），请使用以下打开口令查阅：</p>
<p style="padding:10px 16px;background:#f5f6f8;border-left:3px solid #4a6fa5;border-radius:4px;">
<b>打开口令：{{打开口令}}</b></p>
<p style="color:#888;font-size:12px;">本文件带有贵司专属水印标识，仅供贵司内部使用，请妥善保管、请勿转发。</p>
<p>如有任何疑问，请随时与我们联系。<br><br>此致<br>【公司/团队名称】<br>{{日期}}</p>
</div>"""

# 演示数据（均为虚构，邮箱用保留域名，可直接导入测试全流程）
# 已迁到 lp_v3.py（MOCK_LPS / MOCK_HEADERS），这里保留名字以兼容旧引用。


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# ---------- v2 兼容纯逻辑（迁移、导入与回归测试仍依赖） ----------

def validate_records(lps):
    """批量前统一校验：口令非空/不同、水印非空、跨记录口令不复用。"""
    issues, seen = [], {}
    for r in lps:
        tag = (r.get("name") or "?")[:18]
        if not (r.get("wm") or "").strip():
            issues.append(f"{tag}：水印名称为空（文件将无溯源标识）")
        if not r.get("open_pw"):
            issues.append(f"{tag}：打开口令为空")
        if not r.get("owner_pw"):
            issues.append(f"{tag}：权限口令为空（留档缺失将无法解除限制）")
        elif r.get("open_pw") == r.get("owner_pw"):
            issues.append(f"{tag}：权限口令与打开口令相同（收件人将获得全部权限）")
        if r.get("open_pw"):
            seen.setdefault(r["open_pw"], []).append(tag)
    for pw, names in seen.items():
        if len(names) > 1:
            issues.append("打开口令重复（" + "、".join(names) + "）：同一口令可打开对方文件")
    return issues


def render_filename(tpl, src_path, rec):
    """按模板生成输出文件名（不含扩展名）。字段：{{原件名}} {{LP名称}} {{水印名称}} {{日期}}"""
    mapping = {
        "{{原件名}}": Path(src_path).stem,
        "{{LP名称}}": sanitize_filename(rec.get("name", ""), 60),   # 兼容旧模板
        "{{收件人名称}}": sanitize_filename(rec.get("name", ""), 60),
        "{{水印名称}}": sanitize_filename(rec.get("wm", ""), 60),
        "{{日期}}": date.today().strftime("%Y%m%d"),
    }
    out = tpl or DEFAULT_FILENAME_TPL
    for k, v in mapping.items():
        out = out.replace(k, v)
    out = sanitize_filename(out.strip(), 90)
    return out or "output"


def render_template(tpl, rec, escape=True):
    """v2 邮件正文渲染：只映射 4 个字段，**永远不含权限口令**。"""
    mapping = {
        "{{LP名称}}": rec.get("name", ""),      # 兼容旧模板
        "{{收件人名称}}": rec.get("name", ""),
        "{{水印名称}}": rec.get("wm", ""),
        "{{打开口令}}": rec.get("open_pw", ""),
        "{{日期}}": date.today().strftime("%Y年%m月%d日"),
    }
    for k, v in mapping.items():
        if escape:
            v = _htmlmod.escape(v, quote=True)  # 口令/名称中的 <>& 不破坏邮件HTML
        tpl = tpl.replace(k, v)
    return tpl


def _normalize_word_html(html):
    """Word 另存的 HTML 会把 {{字段}} 按语言标记拆进多个 span，
    导致占位符替换失效——这里把字段内部夹杂的标签剥掉，让字段复原成连续文本。"""
    def rejoin(m):
        return re.sub(r"<[^>]+>", "", m.group(0))
    return re.sub(r"\{\{[^}]*?\}\}", rejoin, html)


# ---------- 内置富文本编辑器：HTML ↔ tk Text 双向转换 ----------

_TAG_STYLE = {}  # tk 标签名 -> 样式 dict（导出时反查）


def _style_key(st):
    return "st:" + ",".join(f"{k}={v}" for k, v in sorted(st.items()) if v)


def _ensure_tag(tw, st):
    name = _style_key(st)
    if name not in _TAG_STYLE:
        fam = "Microsoft YaHei"
        size = st.get("s") or 12
        styles = []
        if st.get("b"):
            styles.append("bold")
        if st.get("i"):
            styles.append("italic")
        if st.get("u"):
            styles.append("underline")
        cfg = {"font": tuple([fam, size] + styles)}
        if st.get("c"):
            cfg["foreground"] = st["c"]
        if st.get("bg"):
            cfg["background"] = st["bg"]
        if st.get("al") in ("center", "right"):
            cfg["justify"] = st["al"]
        tw.tag_configure(name, **cfg)
        _TAG_STYLE[name] = st
    return name


class _HTML2Runs(_HTMLParser):
    """把 HTML 解析成 (纯文本, [(起,止,样式)])，样式用于生成 tk 标签。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.stack = [{}]

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        st = dict(self.stack[-1])
        if tag in ("b", "strong"):
            st["b"] = 1
        elif tag in ("i", "em"):
            st["i"] = 1
        elif tag == "u":
            st["u"] = 1
        elif tag in ("span", "font", "p", "div", "h1", "h2", "h3", "li"):
            style = a.get("style") or ""
            for kv in style.replace("：", ":").split(";"):
                if ":" not in kv:
                    continue
                k, v = (x.strip() for x in kv.split(":", 1))
                k = k.lower()
                if k == "color" and v.startswith("#"):
                    st["c"] = v
                elif k in ("background", "background-color") and v.startswith("#"):
                    st["bg"] = v
                elif k == "font-size":
                    num = v.replace("px", "").replace("pt", "").strip()
                    try:
                        if float(num) > 0:
                            st["s"] = int(float(num))
                    except ValueError:
                        pass
                elif k == "text-align" and v in ("center", "right"):
                    st["al"] = v
            if tag == "font" and (a.get("color") or "").startswith("#"):
                st["c"] = a["color"]
            if tag == "h1":
                st.update(b=1, s=20)
            elif tag == "h2":
                st.update(b=1, s=17)
            elif tag == "h3":
                st.update(b=1, s=14)
            if a.get("align") in ("center", "right"):
                st["al"] = a["align"]
        if tag in ("p", "div", "h1", "h2", "h3", "li", "br", "tr", "table"):
            # 块界换行：已有内容且上一个不是换行时补一个（连续块界不叠加）。
            # 换行符必须不带样式：tk 中带背景标签的换行符会把色条画到编辑区右缘
            if self.parts and self.parts[-1][0] != "\n":
                self.parts.append(("\n", {}))
        if tag == "br":
            self.parts.append(("\n", {}))  # <br> 恒定一个换行（空行=块界+br）
        if tag == "li":
            self.parts.append(("• ", {}))
        if tag in ("span", "font", "b", "strong", "i", "em", "u",
                   "p", "div", "h1", "h2", "h3", "li"):
            self.stack.append(st)

    def handle_endtag(self, tag):
        if tag in ("span", "font", "b", "strong", "i", "em", "u",
                   "p", "div", "h1", "h2", "h3", "li") and len(self.stack) > 1:
            self.stack.pop()

    def handle_data(self, d):
        if d and d != "\n":
            self.parts.append((d, dict(self.stack[-1])))


def html_to_runs(html_str):
    p = _HTML2Runs()
    p.feed(_normalize_word_html(html_str or ""))
    p.close()
    text = "".join(x[0] for x in p.parts)
    runs, pos = [], 0
    for frag, st in p.parts:
        key = _style_key(st)
        if runs and _style_key(runs[-1][2]) == key:
            runs[-1][1] = pos + len(frag)
        else:
            runs.append([pos, pos + len(frag), dict(st)])
        pos += len(frag)
    return text, runs


def load_html(tw, html_str):
    text, runs = html_to_runs(html_str)
    lead = len(text) - len(text.lstrip("\n"))  # 去掉段首累积出的空行
    text = text[lead:]
    tw.delete("1.0", "end")
    tw.insert("1.0", text)
    for a, b, st in runs:
        a, b = max(0, a - lead), max(0, b - lead)
        if a < b:
            tw.tag_add(_ensure_tag(tw, st), f"1.0+{a}c", f"1.0+{b}c")
    # 对齐标签扩展到整行（tk 的 justify 需要覆盖整行才生效）
    n_lines = int(tw.index("end-1c").split(".")[0])
    for ln in range(1, n_lines + 1):
        al = None
        for i in range(len(tw.get(f"{ln}.0", f"{ln}.end"))):
            al = _style_at(tw, f"{ln}.{i}").get("al")
            if al:
                break
        if al:
            tw.tag_add(_ensure_tag(tw, {"al": al}), f"{ln}.0", f"{ln}.end")


def _style_at(tw, index):
    st = {}
    for t in tw.tag_names(index):
        if t in _TAG_STYLE:
            st.update(_TAG_STYLE[t])
    return st


def text_to_html(tw, base=(14, "#333333")):
    """把编辑器内容序列化为 HTML（逐行成段，等样式合并成 span）。"""
    n_lines = int(tw.index("end-1c").split(".")[0])
    out = []
    for ln in range(1, n_lines + 1):
        raw = tw.get(f"{ln}.0", f"{ln}.end")
        parts = []
        for i in range(len(raw)):
            st = _style_at(tw, f"{ln}.{i}")
            k = _style_key(st)
            if parts and parts[-1][2] == k:
                parts[-1][0] += raw[i]
            else:
                parts.append([raw[i], st, k])
        align = parts[0][1].get("al") if parts else None
        seg = ""
        for txt_, st, _k in parts:
            esc = _htmlmod.escape(txt_)
            if not any(st.get(x) for x in ("b", "i", "u", "c", "s", "bg")):
                seg += esc
                continue
            inner = esc
            if st.get("b"):
                inner = f"<b>{inner}</b>"
            if st.get("i"):
                inner = f"<i>{inner}</i>"
            if st.get("u"):
                inner = f"<u>{inner}</u>"
            style = ""
            if st.get("c") and st["c"].lower() != base[1].lower():
                style += f"color:{st['c']};"
            if st.get("s") and st["s"] != base[0]:
                style += f"font-size:{st['s']}px;"
            if st.get("bg"):
                style += f"background:{st['bg']};"
            if style:
                inner = f'<span style="{style[:-1]}">{inner}</span>'
            seg += inner
        alattr = f' align="{align}"' if align else ""
        out.append(f"<p{alattr}>{seg}</p>" if seg else "<p><br></p>")
    return ('<div style="font-family:微软雅黑,Segoe UI,Arial;'
            f'font-size:{base[0]}px;color:{base[1]};line-height:1.8">'
            + "".join(out) + "</div>")


def clip_html_fragment():
    """读取剪贴板里的 HTML 格式（从 Word/网页复制时存在），失败返回 None。"""
    try:
        import win32clipboard
        fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(fmt):
                return None
            raw = win32clipboard.GetClipboardData(fmt)
            if isinstance(raw, str):
                raw = raw.encode("utf-8", errors="replace")
            head = raw[:512].decode("ascii", errors="ignore")
            ms = re.search(r"StartFragment:(\d+)", head)
            me = re.search(r"EndFragment:(\d+)", head)
            if ms and me:
                s, e = int(ms.group(1)), int(me.group(1))
                frag = raw[s:e].decode("utf-8", errors="replace")
            else:
                frag = raw.decode("utf-8", errors="replace")
            m = re.search(r"<body[^>]*>(.*)</body>", frag, re.S | re.I)
            return (m.group(1) if m else frag).strip() or None
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        return None


def render_preview_pixmap(html, width=620, dpi=110):
    """用 PyMuPDF Story 引擎把 HTML 渲染成 tk 可用的 PNG 字节（近似预览）。"""
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
        story.place(mediabox + (pad, pad, -pad, -pad))
        story.draw(dev)
        writer.end_page()
        writer.close()
        doc = fitz.open(path)
        page = doc[0]
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


# ---------- 悬停提示 ----------

class ToolTip:
    def __init__(self, widget, text, delay=450):
        self.widget, self.text, self.delay = widget, text, delay
        self.tip = None
        self.id = None
        widget.bind("<Enter>", self._enter)
        widget.bind("<Leave>", self._leave)
        widget.bind("<ButtonPress>", self._leave)

    def _enter(self, _e):
        self.id = self.widget.after(self.delay, self._show)

    def _leave(self, _e):
        if self.id:
            self.widget.after_cancel(self.id)
        if self.tip:
            self.tip.destroy()
            self.tip = None

    def _show(self):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 12
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left", background="#ffffe8",
                 relief="solid", borderwidth=1, font=("Microsoft YaHei", 9),
                 padx=8, pady=5, wraplength=380).pack()


# ---------- GUI ----------

WB_STEPS = [
    ("收件人", "recipients"),
    ("文档流水线", "pipeline"),
    ("附件", "attachments"),
    ("邮件模板", "mail"),
    ("总览", "overview"),
    ("预览与生成", "run"),
]

# 步骤索引（写死数字容易改一步就错，用常量）
WB_STEP_RECIPIENTS = 0
WB_STEP_PIPELINE = 1
WB_STEP_ATTACHMENTS = 2
WB_STEP_MAIL = 3
WB_STEP_OVERVIEW = 4
WB_STEP_RUN = 5


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1280x800")   # 默认正常窗口（首次不最大化）
        root.minsize(960, 620)
        _icon = Path(__file__).parent / "icon.ico"
        if _icon.exists():
            try:
                root.iconbitmap(str(_icon))
            except Exception:
                pass

        self._boot_error = None
        self.store = Store()
        self.settings = dict(self.store.load_settings()[0])
        self._run_migration()

        self.cur_event_id = None
        self.cur_event = None
        self.running = False
        self._q = _queue.Queue()
        self._jobs = _queue.Queue()
        self._worker_thread = None
        self._preview_img = None
        self._wm_preview_img = None
        self._screens = {}
        self._wb_step = 0
        self._wb_content = None
        self._wb_step_buttons = []
        self._wb_event = None
        self._wb_step_sub = ["", "", "", "", ""]
        self._poll_id = None
        self._shell_widgets = []
        self._wb_last_run = None

        self._build_shell()
        if getattr(self, "_boot_error", None):
            self.root.after(200, lambda: messagebox.showerror("数据读取失败", self._boot_error))
        self.root.after(150, self._poll_queue)
        if self.settings.get("auto_update", True):
            self.root.after(2500, lambda: self.check_update(auto=True))
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ================= 启动 / 迁移 =================

    def _run_migration(self):
        """首次启动：v2 的 lp_data.json → lp_data/ 目录；没有旧数据则播种示例。"""
        if not needs_migration(self.store):
            self.settings, err = self.store.load_settings()
            if err:
                self._boot_error = err
            # 空目录（第一次装 v3）→ 播种示例数据
            if not self.store.list_scenarios() and not self.store.list_events():
                try:
                    ok, notes = lp_v3.seed_demo(self.store)
                    if ok:
                        self.root.after(300, lambda: messagebox.showinfo(
                            "已载入示例数据",
                            "第一次运行，已为你准备了一份示例数据：\n\n"
                            + "\n".join("· " + n for n in notes)
                            + "\n\n收件人邮箱都是虚构的（保留域名），"
                              "可以放心点「生成全部 Outlook 草稿」试跑全流程。"))
                except Exception as e:
                    self._boot_error = "示例数据初始化失败：%s" % e
            return
        ok, notes, err = migrate_v2_to_v3(self.store)
        self.settings, serr = self.store.load_settings()
        if serr:
            self._boot_error = serr
        if ok:
            self.root.after(300, lambda: messagebox.showinfo(
                "已升级到 v3 数据结构",
                "已把 v2 数据迁移为 lp_data/ 目录结构：\n\n"
                + "\n".join("· " + n for n in notes)
                + "\n\n原 lp_data.json 已改名保留为 lp_data.json.bak（未删除，可回退）。"))
        elif err:
            self._boot_error = "数据迁移未完成：%s" % err

    # ================= 外壳 =================

    def _build_shell(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Nav.TButton", padding=(10, 7), anchor="w")
        style.configure("Step.TButton", padding=(8, 6), anchor="w")
        style.configure("Head.TLabel", font=("", 15, "bold"))
        style.configure("Sub.TLabel", foreground="#666")
        style.configure("Card.TLabelframe", padding=8)

        # --- 顶栏 ---
        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=(8, 4))
        self._shell_widgets.append(top)
        ttk.Label(top, text="文档分发助手", style="Head.TLabel").pack(side="left")
        self.crumb = ttk.Label(top, text="", style="Sub.TLabel")
        self.crumb.pack(side="left", padx=14)
        ttk.Button(top, text="检查更新", command=lambda: self.check_update(False)
                   ).pack(side="right")
        ttk.Label(top, text="v" + APP_VERSION, foreground="#888").pack(side="right", padx=10)

        # --- 主体：左导航 + 右内容 ---
        body = ttk.Frame(self.root)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self._shell_widgets.append(body)
        nav = ttk.Frame(body, width=152)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        ttk.Label(nav, text="主导航", foreground="#888").pack(anchor="w", padx=8, pady=(4, 2))
        for key, text in (("home", "事件库"),):
            b = ttk.Button(nav, text=text, style="Nav.TButton",
                           command=lambda k=key: self.show_screen(k))
            b.pack(fill="x", padx=4, pady=1)
        ttk.Label(nav, text="全局", foreground="#888").pack(anchor="w", padx=8, pady=(12, 2))
        for key, text in (("settings", "设置"), ("help", "帮助")):
            b = ttk.Button(nav, text=text, style="Nav.TButton",
                           command=lambda k=key: self.show_screen(k))
            b.pack(fill="x", padx=4, pady=1)

        self.content = ttk.Frame(body)
        self.content.pack(side="left", fill="both", expand=True)
        # 启动后直接落在「场景与事件」页
        self.show_screen("home")

    def show_screen(self, name):
        """切换顶层页面（首次进入时构建）。"""
        for f in self._screens.values():
            f.pack_forget()
        if name not in self._screens:
            frame = ttk.Frame(self.content)
            self._screens[name] = frame
            builder = {
                "home": self._build_home,
                "settings": self._build_settings,
                "help": self._build_help,
                "event": self._build_workbench,
            }.get(name)
            if builder:
                builder(frame)
        self._screens[name].pack(fill="both", expand=True)
        self._set_crumb({
            "home": "事件库", "settings": "设置", "help": "帮助",
        }.get(name, name))
        if name == "home":
            self._refresh_home()
        elif name == "settings":
            self._refresh_settings()

    def _set_crumb(self, text):
        self.crumb.configure(text=text)

    def _guide(self, parent, text):
        """页顶部的说明条。"""
        ttk.Label(parent, text=text, background="#eef3fa", foreground="#2b4a75",
                  padding=(10, 7), wraplength=1080, justify="left").pack(
            fill="x", pady=(0, 4))

    # ================= 事件库（主页面） =================

    def _build_home(self, parent=None):
        """事件库：唯一的主页面。

        一次分发 = 一个事件，里面装「发送模式 + 收件人 + 流水线 + 邮件模板」。
        「场景模板」收进同一页的第二个标签页，不再单独占一个导航入口
        （之前是「场景与事件 / 场景模板 / 收件人库」三个入口，用户看不懂区别）。
        """
        f = parent if parent is not None else self._screens["home"]

        head = ttk.Frame(f)
        head.pack(fill="x", pady=(0, 4))
        ttk.Label(head, text="事件库", style="Head.TLabel").pack(side="left")
        ttk.Label(head, text="一次分发 = 一个事件：发送模式 + 收件人 + 流水线 + 邮件模板",
                  style="Sub.TLabel").pack(side="left", padx=12)

        self.home_nb = ttk.Notebook(f)
        self.home_nb.pack(fill="both", expand=True)

        self.home_tab_ev = ttk.Frame(self.home_nb)
        self.home_nb.add(self.home_tab_ev, text=" 事件 ")
        self.home_tab_sc = ttk.Frame(self.home_nb)
        self.home_nb.add(self.home_tab_sc, text=" 场景模板 ")
        # P4：收件人库与执行历史做成标签页，不占顶层导航入口
        self.home_tab_pool = ttk.Frame(self.home_nb)
        self.home_nb.add(self.home_tab_pool, text=" 收件人库 ")
        self.home_tab_runs = ttk.Frame(self.home_nb)
        self.home_nb.add(self.home_tab_runs, text=" 执行历史 ")

        self._build_events_tab(self.home_tab_ev)
        self._build_scenes(self.home_tab_sc)
        self._build_pool_tab(self.home_tab_pool)
        self._build_runs_tab(self.home_tab_runs)

    def _build_events_tab(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 4))
        for text, cmd in (("新建事件", self._new_event),
                          ("复制", self._dup_event),
                          ("删除", self._del_event),
                          ("打开工作台", self._open_selected_event)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=(0, 4))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        for text, cmd in (("导入收件人…", self._lib_import_recipients),
                          ("导出收件人…", self._lib_export_recipients),
                          ("导出全部收件人…", self._lib_export_all_recipients),
                          ("导出事件…", self._lib_export_event),
                          ("导入事件备份…", self._lib_import_event),
                          ("备份全部…", self._lib_export_backup)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=(0, 4))
        ttk.Button(bar, text="刷新", command=self._refresh_home).pack(side="left", padx=8)
        self.home_status = ttk.Label(bar, text="", style="Sub.TLabel")
        self.home_status.pack(side="right")

        paned = tk.PanedWindow(parent, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # 左：事件表
        left = ttk.Frame(paned)
        cols = ["name", "mode", "recips", "steps", "atts", "status", "last"]
        heads = ["事件名称", "发送模式", "收件人", "流水线", "附件", "状态", "最近执行"]
        self.ev_tree = ttk.Treeview(left, columns=cols, show="headings",
                                    selectmode="browse")
        for c, h in zip(cols, heads):
            self.ev_tree.heading(c, text=h)
            self.ev_tree.column(c, width=190 if c == "name" else 86)
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.ev_tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.ev_tree.xview)
        self.ev_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.ev_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.ev_tree.bind("<Double-1>", lambda _e: self._open_selected_event())
        self.ev_tree.bind("<<TreeviewSelect>>", lambda _e: self._draw_selected_flow())

        # 右：流程图
        right = ttk.Frame(paned)
        self.flow_frame = ttk.LabelFrame(right, text=" 流程图 ",
                                         style="Card.TLabelframe")
        self.flow_frame.pack(fill="both", expand=True)
        self.flow_title = ttk.Label(self.flow_frame, text="", style="Sub.TLabel",
                                    wraplength=460, justify="left")
        self.flow_title.pack(anchor="w", padx=8, pady=(4, 0))
        self.flow_canvas = tk.Canvas(self.flow_frame, height=300,
                                     highlightthickness=0, background="#fbfcfe")
        self.flow_canvas.pack(fill="both", expand=True, padx=6, pady=6)
        self.flow_canvas.bind("<Configure>", lambda _e: self._draw_selected_flow())

        paned.add(left)
        paned.add(right)
        # weight 让左右两栏按 3:2 分配（部分 Tk 版本 add 不接受 weight，失败不影响使用）
        for w, child in ((3, left), (2, right)):
            try:
                paned.paneconfigure(child, weight=w)
            except Exception:
                pass

    def _refresh_home(self):
        self._refresh_events_tab()
        self._refresh_scenes()
        self._refresh_pool_tab()
        self._refresh_runs_tab()

    def _refresh_events_tab(self):
        events = self.store.list_events()
        self.ev_tree.delete(*self.ev_tree.get_children())
        for e in events:
            last = ""
            st = e.get("status", "")
            runs = self.store.list_runs(e["id"])
            if runs:
                r = runs[0]
                items = r.get("items") or []
                ok = sum(1 for i in items if i.get("draft") == "ok")
                last = "%s · %d/%d" % ((r.get("started_at") or "")[5:16], ok, len(items))
                if not st:
                    st = "已生成草稿"
            mode = lp_v3.send_mode_of(e)
            self.ev_tree.insert("", "end", iid=e["id"], values=[
                e.get("name", ""),
                lp_v3.SEND_MODE_LABELS.get(mode, mode),
                len(e.get("recipients", [])), len(e.get("pipeline", [])),
                len(e.get("attachments", [])), st or "草稿", last or "—"])
        self.home_status.configure(text="共 %d 个事件" % len(events))
        self._draw_selected_flow()

    # ---------- 流程图 ----------

    def _draw_selected_flow(self):
        eid = self._selected_event_id()
        ev = None
        if eid:
            ev, _err = self.store.load_event(eid)
        if ev:
            self.flow_title.configure(
                text="事件「%s」· %s · 收件人 %d 位 · 流水线 %d 个条目"
                     % (ev.get("name", ""),
                        lp_v3.SEND_MODE_LABELS.get(lp_v3.send_mode_of(ev), ""),
                        len(ev.get("recipients", [])), len(ev.get("pipeline", []))))
        else:
            self.flow_title.configure(
                text="在上表里选中一个事件，这里会画出它的流程图"
                     "（条目可在「文档流水线」里自由增删排序）")
        self._draw_flow(self.flow_canvas, ev)

    def _draw_flow(self, canvas, event):
        """把事件画成流程图：收件人 → 流水线条目 → 附件包 → Outlook 草稿。"""
        canvas.delete("all")
        if event is None:
            canvas.create_text(
                16, 16, anchor="nw", fill="#888888", justify="left",
                text="还没有选中事件。\n\n"
                     "流程图按「收件人 → 流水线条目 → 附件包 → Outlook 草稿」画出来，\n"
                     "条目可以在「文档流水线」里自由增删和排序。")
            return
        nodes = lp_v3.flow_nodes(event)
        width = max(canvas.winfo_width(), 360)
        box_h, gap = 56, 26
        box_w = 152
        avail = width - 24
        # 优先把整条链压成一行（盒子太窄就换行，避免挤成一团）
        if len(nodes) * (box_w + gap) > avail:
            fit = int(avail / len(nodes)) - gap
            if fit >= 96:
                box_w = fit
            else:
                box_w = 152
        per_row = max(1, (avail) // (box_w + gap))
        if len(nodes) * (box_w + gap) <= avail:
            per_row = len(nodes)
        rows = (len(nodes) + per_row - 1) // per_row
        need_h = rows * (box_h + 34) + 12
        if canvas.winfo_height() < need_h:
            canvas.configure(height=need_h)
        styles = {
            "source": ("#eef3fa", "#4a6fa5", "#1f3d63"),
            "step":   ("#ffffff", "#4a6fa5", "#1f3d63"),
            "output": ("#eaf7ea", "#3f8f4f", "#1e5227"),
            "draft":  ("#fdf3e3", "#c08a2e", "#6b4712"),
        }
        pos = {}
        for i, n in enumerate(nodes):
            r, c = divmod(i, per_row)
            x = 14 + c * (box_w + gap)
            y = 12 + r * (box_h + 34)
            pos[i] = (x, y)
            fill, outline, fg = styles.get(n.get("kind"), styles["step"])
            if not n.get("enabled", True):
                fill, outline, fg = "#f3f3f3", "#bcbcbc", "#8a8a8a"
            canvas.create_rectangle(x, y, x + box_w, y + box_h,
                                    fill=fill, outline=outline, width=2)
            canvas.create_text(x + box_w / 2, y + 20, text=n.get("title", ""),
                               font=("", 9, "bold"), fill=fg, width=box_w - 14)
            canvas.create_text(x + box_w / 2, y + 39, text=n.get("sub", ""),
                               font=("", 8), fill="#666666", width=box_w - 14)
        for i in range(len(nodes) - 1):
            x0, y0 = pos[i]
            x1, y1 = pos[i + 1]
            if (i // per_row) == ((i + 1) // per_row):
                canvas.create_line(x0 + box_w, y0 + box_h / 2, x1, y1 + box_h / 2,
                                   fill="#8ea6c4", width=2, arrow="last")
            else:
                mid = y0 + box_h + 17
                canvas.create_line(x0 + box_w / 2, y0 + box_h,
                                   x0 + box_w / 2, mid, x1, mid, x1, y1,
                                   fill="#8ea6c4", width=2, arrow="last")
        canvas.configure(scrollregion=(0, 0, width, need_h))


    def _new_scenario(self, from_event=None):
        name = simpledialog.askstring("新建场景", "场景名称：", initialvalue="新建场景")
        if not name:
            return
        sc = new_scenario(name)
        if from_event:
            sc["recipients"] = [normalize_recipient(r) for r in from_event.get("recipients", [])]
            sc["pipeline"] = [normalize_item(s) for s in from_event.get("pipeline", [])]
            sc["attachments"] = [normalize_attachment(a) for a in from_event.get("attachments", [])]
            sc["mail"] = dict(from_event.get("mail") or {})
            sc["options"] = dict(from_event.get("options") or {})
        ok, err = self.store.save_scenario(sc)
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        self._refresh_home()

    def _new_event(self, scenario=None):
        ev = new_event("未命名事件", scenario)
        ok, err = self.store.save_event(ev)
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        self._refresh_home()
        self.open_event(ev["id"])

    def _selected_event_id(self):
        sel = self.ev_tree.selection()
        return sel[0] if sel else None

    def _dup_event(self):
        eid = self._selected_event_id()
        if not eid:
            messagebox.showinfo("提示", "请先选中一个事件")
            return
        ev, err = self.store.load_event(eid)
        if not ev:
            messagebox.showerror("读取失败", err or "事件不存在")
            return
        ok, err = self.store.save_event(clone_event(ev))
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        self._refresh_home()

    def _del_event(self):
        eid = self._selected_event_id()
        if not eid:
            messagebox.showinfo("提示", "请先选中一个事件")
            return
        ev, _err = self.store.load_event(eid)
        if not messagebox.askyesno("删除事件",
                                   "确认删除事件「%s」？\n\n（执行记录会保留）"
                                   % (ev.get("name", "") if ev else eid)):
            return
        self.store.delete_event(eid)
        self._refresh_home()

    def _open_selected_event(self):
        eid = self._selected_event_id()
        if not eid:
            messagebox.showinfo("提示", "请先选中一个事件")
            return
        self.open_event(eid)

    # ================= 收件人库（事件库里的第三个标签页，P4） =================
    #
    # 为什么不做成独立导航页：用户 2026-10-23 明确说过「场景与事件 / 场景模板 / 收件人库
    # 三个入口看不懂区别，只要事件库」，当时把收件人库那个空壳页删了。
    # 所以这里做进事件库的标签页，顶层导航仍然是「事件库 / 设置 / 帮助」。

    def _build_pool_tab(self, parent=None):
        f = parent if parent is not None else getattr(self, "home_tab_pool", None)
        if f is None:
            return
        ttk.Label(f, text="收件人库", style="Head.TLabel").pack(anchor="w")
        ttk.Label(f, text=(
            "常用的那批收件人，跨事件复用。\n"
            "「从事件导入…」把某个事件的收件人整个收进库；「添加到事件…」把库里的人并进某个事件。\n"
            "合并按收件人名称进行；「事件里已有的打开口令 / 权限口令」不会被覆盖。\n"
            "库里也存着口令，与 lp_data/ 其余部分同管控级别。"),
            style="Sub.TLabel", justify="left").pack(anchor="w", pady=(0, 6))

        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(0, 4))
        for text, cmd, tip in (
                ("新增", self._pool_new, "在收件人库里新增一条"),
                ("编辑", self._pool_edit, "编辑选中的收件人"),
                ("删除", self._pool_delete, "从库里删除选中的收件人"),
                ("从事件导入…", self._pool_from_event, "把某个事件的收件人整个收进库"),
                ("添加到事件…", self._pool_to_event, "把选中的收件人并进某个事件"),
                ("导入 Excel/CSV", self._pool_import, "按收件人清单模板导入（可先下载模板）"),
                ("导出 Excel", self._pool_export, "把收件人库导出成 Excel"),
        ):
            b = ttk.Button(bar, text=text, command=cmd)
            b.pack(side="left", padx=2)
            ToolTip(b, tip)
        b = ttk.Button(bar, text="下载导入模板", command=self._wb_export_template)
        b.pack(side="left", padx=(10, 2))
        ToolTip(b, "收件人清单模板（含填写说明页），与事件里用的是同一份")

        frame = ttk.Frame(f)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=["name"], show="headings",
                            selectmode="extended")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.bind("<Double-1>", lambda _e: self._pool_edit())
        self._pool_tree = tree
        self._pool_iids = {}

        ttk.Label(f, foreground="#b00", text=(
            "安全提示：收件人库里含打开口令与权限口令，属敏感文件；"
            "本程序同目录的 lp_data/ 同理，请限定在本机/受控共享目录使用。")
        ).pack(fill="x", pady=(4, 0))

    def _refresh_pool_tab(self):
        tree = getattr(self, "_pool_tree", None)
        if tree is None:
            return
        pool = self.store.list_pool()
        fcols = []
        for r in pool:
            for k in (r.get("fields") or {}):
                if k not in fcols:
                    fcols.append(k)
        cols = ["name", "to", "cc", "bcc", "wm", "group"] + \
            ["f%d" % i for i in range(len(fcols))] + ["open_pw", "owner_pw"]
        heads = ["收件人名称", "To", "CC", "BCC", "水印", "组"] + fcols + \
            ["打开口令", "权限口令(勿发)"]
        tree.delete(*tree.get_children())
        self._pool_iids = {}
        try:
            tree.configure(columns=cols)
        except Exception:
            pass
        for c, h in zip(cols, heads):
            try:
                tree.heading(c, text=h)
                tree.column(c, width=150 if c in ("to", "cc", "bcc") else 110)
            except Exception:
                pass
        for i, r in enumerate(pool):
            vals = [r.get("name", ""), r.get("to", ""), r.get("cc", ""),
                    r.get("bcc", ""), r.get("wm", ""), r.get("group", "")]
            vals += [(r.get("fields") or {}).get(k, "") for k in fcols]
            vals += [r.get("open_pw", ""),
                     "••••" if r.get("owner_pw") else ""]
            iid = "p%d" % i
            self._pool_iids[iid] = r.get("id", "")
            tree.insert("", "end", iid=iid, values=vals)

    def _pool_selected(self):
        tree = getattr(self, "_pool_tree", None)
        if tree is None:
            return []
        ids = self._pool_iids or {}
        return [ids[i] for i in tree.selection() if i in ids]

    def _pool_new(self):
        rec = lp_v3.new_pool_recipient({
            "name": "", "wm": "", "to": "", "cc": "", "bcc": "",
            "open_pw": gen_pw(), "owner_pw": gen_pw(), "fields": {},
            "attachments": [], "group": ""})

        def save(out):
            out["id"] = rec.get("id", "")
            ok, err, _rec = self.store.save_pool_recipient(out)
            if not ok:
                messagebox.showerror("保存失败", err)
                return False
            self._refresh_pool_tab()
            return True

        self._recipient_dialog(rec, self._pool_field_cols(),
                               "新增收件人（收件人库）", save,
                               "（库里还没有自定义字段列）")

    def _pool_field_cols(self):
        cols = []
        for r in self.store.list_pool():
            for k in (r.get("fields") or {}):
                if k not in cols:
                    cols.append(k)
        return cols

    def _pool_edit(self):
        tree = getattr(self, "_pool_tree", None)
        sel = tree.selection() if tree else ()
        if not sel:
            messagebox.showinfo("提示", "请先选中一位收件人")
            return
        rid = (self._pool_iids or {}).get(sel[0])
        rec = None
        for r in self.store.list_pool():
            if r.get("id") == rid:
                rec = dict(r)
                break
        if rec is None:
            return

        def save(out):
            out["id"] = rid
            ok, err, _rec = self.store.save_pool_recipient(out)
            if not ok:
                messagebox.showerror("保存失败", err)
                return False
            self._refresh_pool_tab()
            return True

        self._recipient_dialog(rec, self._pool_field_cols(),
                               "编辑收件人（收件人库）", save,
                               "（库里还没有自定义字段列）")

    def _pool_delete(self):
        ids = self._pool_selected()
        if not ids:
            messagebox.showinfo("提示", "请先选中要删除的收件人")
            return
        if not messagebox.askyesno("确认删除",
                                   "确定要从收件人库里删除 %d 位收件人吗？\n"
                                   "（不会影响任何事件里的收件人）" % len(ids)):
            return
        for rid in ids:
            ok, err = self.store.delete_pool_recipient(rid)
            if not ok:
                messagebox.showerror("删除失败", err)
                return
        self._refresh_pool_tab()

    def _pick_event(self, title):
        """弹窗选一个事件。返回 event id 或 None。"""
        events = self.store.list_events()
        if not events:
            messagebox.showinfo("提示", "还没有任何事件")
            return None
        win = tk.Toplevel(self.root)
        win.title(title)
        win.grab_set()
        win.geometry("460x360")
        ttk.Label(win, text=title, justify="left").pack(anchor="w", padx=12, pady=(10, 4))
        lb = tk.Listbox(win, exportselection=False)
        lb.pack(fill="both", expand=True, padx=12)
        for e in events:
            lb.insert("end", "%s（%d 位收件人）" % (e.get("name", ""),
                                                  len(e.get("recipients", []))))
        picked = {"id": None}

        def ok():
            sel = lb.curselection()
            if sel:
                picked["id"] = events[sel[0]].get("id")
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(pady=8)
        ttk.Button(btns, text="确定", command=ok).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=6)
        return picked["id"]

    def _pool_from_event(self):
        eid = self._pick_event("把哪个事件的收件人收进收件人库？")
        if not eid:
            return
        ev, _err = self.store.load_event(eid)
        if not ev:
            return
        recips = ev.get("recipients") or []
        if not recips:
            messagebox.showinfo("提示", "这个事件还没有收件人")
            return
        pool = self.store.list_pool()
        added = updated = 0
        for rec in recips:
            src = dict(rec)
            src.pop("id", None)
            pool, new_rec, is_update = lp_v3.pool_merge_recipient(pool, src)
            if new_rec is not None:
                ok, err, _saved = self.store.save_pool_recipient(new_rec)
                if not ok:
                    messagebox.showerror("保存失败", err)
                    return
                added += 1
            elif is_update:
                updated += 1
        self._refresh_pool_tab()
        messagebox.showinfo("已收入收件人库",
                            "新增 %d 位，按名称合并 %d 位（库里的口令保留）。" % (added, updated))

    def _pool_to_event(self):
        ids = self._pool_selected()
        if not ids:
            messagebox.showinfo("提示", "请先选中要添加的收件人")
            return
        eid = self._pick_event("把选中的收件人添加到哪个事件？")
        if not eid:
            return
        ev, _err = self.store.load_event(eid)
        if not ev:
            return
        pool = {r.get("id"): r for r in self.store.list_pool()}
        recs = [pool[i] for i in ids if i in pool]
        if not recs:
            return
        new_ev, added, updated = lp_v3.merge_recipients_into_event(ev, recs)
        ok, err = self.store.save_event(new_ev)
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        self._refresh_home()
        messagebox.showinfo(
            "已添加",
            "事件「%s」新增 %d 位，按名称合并 %d 位。\n"
            "（事件里已有的打开/权限口令没有被覆盖）"
            % (new_ev.get("name", ""), added, updated))

    def _pool_import(self):
        path = filedialog.askopenfilename(title="选择收件人清单",
                                          filetypes=[("表格", "*.xlsx *.csv"),
                                                     ("所有文件", "*.*")])
        if not path:
            return
        try:
            rows = lp_v3.read_spreadsheet(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        result = lp_v3.validate_recipient_rows(rows)
        action = self._show_import_result(result, path)
        if not action:
            return
        pool = self.store.list_pool()
        added = updated = 0
        for item in result.get("items") or []:
            if action == "ok_only" and item.get("status") == "error":
                continue
            rec = item.get("record") or {}
            if not rec.get("name"):
                continue
            pool, new_rec, is_update = lp_v3.pool_merge_recipient(pool, rec)
            if new_rec is not None:
                ok, err, _saved = self.store.save_pool_recipient(new_rec)
                if not ok:
                    messagebox.showerror("保存失败", err)
                    return
                added += 1
            elif is_update:
                updated += 1
        self._refresh_pool_tab()
        messagebox.showinfo("导入完成",
                            "收件人库新增 %d 位，按名称合并 %d 位。" % (added, updated))

    def _pool_export(self):
        pool = self.store.list_pool()
        if not pool:
            messagebox.showinfo("提示", "收件人库还是空的")
            return
        rows = lp_v3.recipients_to_rows(pool)
        self._write_xlsx_sheets(
            "收件人库.xlsx",
            [("收件人库", rows, [36, 20, 28, 16, 16, 22, 34, 14, 14, 24])],
            "已导出",
            "收件人库已导出。\n\n注意：文件里含打开口令与权限口令，属敏感文件，"
            "请按与 lp_data/ 相同的管控要求保存。")

    # ================= 执行历史（事件库里的第四个标签页，P4） =================

    RUN_STATUS_LABELS = {
        "ok": "已完成",
        "preflight_failed": "预检未通过",
        "failed": "失败",
        "": "未完成",
    }

    def _build_runs_tab(self, parent=None):
        f = parent if parent is not None else getattr(self, "home_tab_runs", None)
        if f is None:
            return
        ttk.Label(f, text="执行历史", style="Head.TLabel").pack(anchor="w")
        ttk.Label(f, text=(
            "每一次生成草稿都会留一条记录（时间、每个收件人拿到了哪些文件、有没有出错）。\n"
            "双击一行看详情。执行记录是审计材料，这里只看不删。"),
            style="Sub.TLabel", justify="left").pack(anchor="w", pady=(0, 6))

        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(0, 4))
        ttk.Button(bar, text="查看详情", command=self._runs_detail).pack(side="left")
        ttk.Button(bar, text="打开该事件", command=self._runs_open_event).pack(side="left", padx=4)
        ttk.Button(bar, text="刷新", command=self._refresh_runs_tab).pack(side="left")

        frame = ttk.Frame(f)
        frame.pack(fill="both", expand=True)
        cols = ["started", "event", "status", "drafts", "ok", "fail", "skip", "files"]
        heads = ["开始时间", "事件", "状态", "草稿", "成功", "失败", "跳过", "附件包"]
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
            tree.column(c, width=150 if c == "started" else (200 if c == "event" else 84))
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.bind("<Double-1>", lambda _e: self._runs_detail())
        self._runs_tree = tree
        self._runs_ids = {}

    def _refresh_runs_tab(self):
        tree = getattr(self, "_runs_tree", None)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        self._runs_ids = {}
        runs = self.store.list_runs()
        for i, r in enumerate(runs):
            items = r.get("items") or []
            drafts = len([x for x in items if (x.get("draft") or "").strip()])
            files = sum(len(x.get("files") or []) for x in items)
            st = r.get("status") or ""
            vals = [
                r.get("started_at", ""),
                r.get("event_name", "") or "（事件已删除）",
                self.RUN_STATUS_LABELS.get(st, st or "未完成"),
                "%d/%d" % (drafts, len(items)),
                r.get("stats", {}).get("ok", 0),
                r.get("stats", {}).get("fail", 0),
                r.get("stats", {}).get("skip", 0),
                files,
            ]
            iid = "r%d" % i
            self._runs_ids[iid] = r
            tree.insert("", "end", iid=iid, values=vals)

    def _runs_selected(self):
        tree = getattr(self, "_runs_tree", None)
        if tree is None:
            return None
        sel = tree.selection()
        if not sel:
            return None
        return (self._runs_ids or {}).get(sel[0])

    def _runs_detail(self):
        run = self._runs_selected()
        if not run:
            messagebox.showinfo("提示", "请先选中一条执行记录")
            return
        win = tk.Toplevel(self.root)
        win.title("执行详情 · %s · %s" % (run.get("event_name", ""),
                                         run.get("started_at", "")))
        win.geometry("1000x620")
        win.transient(self.root)

        head = ttk.Frame(win)
        head.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(head, text="状态：%s" % self.RUN_STATUS_LABELS.get(
            run.get("status") or "", run.get("status") or "未完成"),
            style="Head.TLabel").pack(side="left")
        ttk.Label(head, text="开始 %s" % run.get("started_at", ""),
                  style="Sub.TLabel").pack(side="left", padx=12)
        if run.get("finished_at"):
            ttk.Label(head, text="结束 %s" % run.get("finished_at", ""),
                      style="Sub.TLabel").pack(side="left", padx=12)

        cols = ["who", "members", "files", "missing", "draft", "error"]
        heads = ["收件人/组", "成员", "附件包", "缺失", "草稿", "错误"]
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=4)
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=12)
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
        tree.column("who", width=180)
        tree.column("members", width=220)
        tree.column("files", width=70, stretch=False)
        tree.column("missing", width=200)
        tree.column("draft", width=110, stretch=False)
        tree.column("error", width=260)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        for it in run.get("items") or []:
            members = it.get("members") or []
            tree.insert("", "end", values=(
                it.get("name", ""),
                "、".join(members),
                len(it.get("files") or []),
                "、".join(it.get("missing") or [])[:120],
                it.get("draft", "") or "—",
                (it.get("error") or "")[:160]))

        log_box = ttk.LabelFrame(win, text=" 执行日志 / 错误与提醒 ", style="Card.TLabelframe")
        log_box.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        txt = tk.Text(log_box, height=9, wrap="word")
        txt.pack(fill="both", expand=True, padx=6, pady=6)
        lines = []
        for e in run.get("errors") or []:
            lines.append("【错误】%s" % e)
        for w in run.get("warnings") or []:
            lines.append("【提醒】%s" % w)
        lines.append("")
        lines += list(run.get("logs") or [])
        txt.insert("1.0", "\n".join(lines) or "（无日志）")
        txt.config(state="disabled")

        btns = ttk.Frame(win)
        btns.pack(pady=(0, 10))
        ttk.Button(btns, text="打开该事件",
                   command=lambda: (win.destroy(),
                                    self._runs_open_event())).pack(side="left", padx=6)
        ttk.Button(btns, text="关闭", command=win.destroy).pack(side="left", padx=6)

    def _runs_open_event(self):
        run = self._runs_selected()
        if not run:
            messagebox.showinfo("提示", "请先选中一条执行记录")
            return
        eid = run.get("event_id")
        ev, err = self.store.load_event(eid)
        if not ev:
            messagebox.showinfo("提示", "这个事件已经不存在了")
            return
        self.show_screen("home")
        self.open_event(ev["id"])

    # ================= 场景模板（事件库里的第二个标签页） =================

    def _build_scenes(self, parent=None):
        f = parent if parent is not None else getattr(self, "home_tab_sc", None)
        if f is None:
            return
        head = ttk.Frame(f)
        head.pack(fill="x")
        ttk.Label(head, text="场景模板", style="Head.TLabel").pack(side="left")
        ttk.Label(head, text="可复用的「怎么做」模板。场景修改不会影响已创建的事件"
                             "——事件在创建时复制了一份配置。",
                  style="Sub.TLabel").pack(side="left", padx=12)
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=4)
        ttk.Button(bar, text="新建场景", command=self._new_scenario).pack(side="left")
        ttk.Button(bar, text="重命名", command=self._rename_scenario).pack(side="left", padx=4)
        ttk.Button(bar, text="删除场景", command=self._del_scenario).pack(side="left")
        ttk.Button(bar, text="以此新建事件", command=self._scene_new_event).pack(side="left", padx=8)
        ttk.Button(bar, text="同步到事件…", command=self._scene_sync_event,
                   ).pack(side="left")
        ToolTip(bar.winfo_children()[-1],
                "把当前场景的流水线/附件/邮件模板/选项覆盖到选中的事件。\n"
                "收件人按名称合并：已存在的收件人保留自己的口令与自定义字段，"
                "只补齐空缺。\n事件不会自动跟随场景变化，需要你显式点这个按钮。")
        cols = ["name", "desc", "recips", "steps", "atts", "derived", "go"]
        heads = ["名称", "说明", "收件人模板", "流水线", "附件", "派生事件", ""]
        self.sc_tree = ttk.Treeview(f, columns=cols, show="headings", selectmode="browse")
        for c, h in zip(cols, heads):
            self.sc_tree.heading(c, text=h)
            self.sc_tree.column(c, width=200 if c in ("name", "desc") else 90)
        vsb = ttk.Scrollbar(f, orient="vertical", command=self.sc_tree.yview)
        self.sc_tree.configure(yscrollcommand=vsb.set)
        self.sc_tree.pack(side="top", fill="both", expand=True, padx=4)
        vsb.pack(side="right", fill="y")

    def _refresh_scenes(self):
        self.sc_tree.delete(*self.sc_tree.get_children())
        events = self.store.list_events()
        derived = {}
        for e in events:
            derived[e.get("scenario_id", "")] = derived.get(e.get("scenario_id", ""), 0) + 1
        for sc in self.store.list_scenarios():
            self.sc_tree.insert("", "end", iid=sc["id"], values=[
                sc.get("name", ""), sc.get("description", ""),
                len(sc.get("recipients", [])), len(sc.get("pipeline", [])),
                len(sc.get("attachments", [])), derived.get(sc["id"], 0), "›"])

    def _selected_scenario_id(self):
        sel = self.sc_tree.selection()
        return sel[0] if sel else None

    def _rename_scenario(self):
        sid = self._selected_scenario_id()
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景")
            return
        sc, err = self.store.load_scenario(sid)
        if not sc:
            messagebox.showerror("读取失败", err or "场景不存在")
            return
        name = simpledialog.askstring("重命名场景", "场景名称：", initialvalue=sc.get("name", ""))
        if not name:
            return
        sc["name"] = name
        self.store.save_scenario(sc)
        self._refresh_scenes()

    def _del_scenario(self):
        sid = self._selected_scenario_id()
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景")
            return
        sc, _err = self.store.load_scenario(sid)
        if not messagebox.askyesno("删除场景",
                                   "确认删除场景「%s」？\n\n已创建的事件不受影响。"
                                   % (sc.get("name", "") if sc else sid)):
            return
        self.store.delete_scenario(sid)
        self._refresh_scenes()

    def _scene_new_event(self):
        sid = self._selected_scenario_id()
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景")
            return
        sc, err = self.store.load_scenario(sid)
        if not sc:
            messagebox.showerror("读取失败", err or "场景不存在")
            return
        self._new_event(sc)

    def _scene_sync_event(self):
        """手动把场景配置同步到指定事件（复制式，用户显式触发）。"""
        sid = self._selected_scenario_id()
        if not sid:
            messagebox.showinfo("提示", "请先选中一个场景")
            return
        sc, err = self.store.load_scenario(sid)
        if not sc:
            messagebox.showerror("读取失败", err or "场景不存在")
            return
        events = self.store.list_events()
        if not events:
            messagebox.showinfo("提示", "还没有发送事件")
            return
        names = [e.get("name", "") for e in events]
        pick = simpledialog.askstring(
            "同步到事件",
            "要把场景「%s」同步到哪个事件？\n请输入事件名称：\n\n%s"
            % (sc.get("name", ""), "\n".join(names)), initialvalue=names[0])
        if not pick:
            return
        target = next((e for e in events if e.get("name") == pick.strip()), None)
        if target is None:
            messagebox.showwarning("同步", "没有找到名为「%s」的事件" % pick)
            return
        if not messagebox.askyesno(
                "确认同步",
                "将用场景「%s」覆盖事件「%s」的：\n"
                "· 流水线（%d 步）\n· 静态附件（%d 个）\n· 邮件模板与选项\n\n"
                "收件人按名称合并，已有口令与自定义字段保留。\n\n确认同步？"
                % (sc.get("name", ""), target.get("name", ""),
                   len(sc.get("pipeline", [])), len(sc.get("attachments", [])))):
            return
        new_ev, notes = sync_scenario_to_event(sc, target)
        ok, err = self.store.save_event(new_ev)
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        messagebox.showinfo("同步完成", "\n".join("· " + n for n in notes))
        self._refresh_scenes()

    # ================= 设置 =================

    def _build_settings(self, parent=None):
        f = parent if parent is not None else self._screens["settings"]
        ttk.Label(f, text="设置", style="Head.TLabel").pack(anchor="w", pady=(0, 4))
        ttk.Label(f, text="全局默认：新建事件时带出，可单独覆盖。",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 6))

        row = ttk.Frame(f)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="默认输出目录：").pack(side="left")
        self.set_outdir = tk.StringVar(value=self.settings.get("outdir", ""))
        ttk.Entry(row, textvariable=self.set_outdir, width=52).pack(side="left", padx=4)
        ttk.Button(row, text="浏览", command=self._set_pick_outdir).pack(side="left")
        ttk.Button(row, text="保存", command=self._save_settings).pack(side="left", padx=8)

        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="文件名模板：").pack(side="left")
        self.set_fname = tk.StringVar(value=self.settings.get("filename_tpl", DEFAULT_FILENAME_TPL))
        ttk.Entry(row2, textvariable=self.set_fname, width=40).pack(side="left", padx=4)
        ttk.Label(row2, text="字段：{{原件名}} {{LP名称}} {{水印名称}} {{日期}}",
                  style="Sub.TLabel").pack(side="left", padx=6)
        ttk.Button(row2, text="保存", command=self._save_settings).pack(side="left", padx=8)

        row3 = ttk.Frame(f)
        row3.pack(fill="x", pady=2)
        self.set_gp_mode = tk.BooleanVar(value=bool(self.settings.get("gp_owner_mode")))
        ttk.Checkbutton(row3, text="GP 统一权限口令（每个文件同一个 owner 口令）",
                        variable=self.set_gp_mode).pack(side="left")
        self.set_gp_pw = tk.StringVar(value=self.settings.get("gp_owner_pw", ""))
        ttk.Entry(row3, textvariable=self.set_gp_pw, width=22, show="*").pack(side="left", padx=6)
        ttk.Button(row3, text="保存", command=self._save_settings).pack(side="left", padx=8)

        wm_row = ttk.Frame(f)
        wm_row.pack(fill="x", pady=6)
        ttk.Button(wm_row, text="水印设置…", command=self.open_wm_settings).pack(side="left")
        self.wm_summary = ttk.Label(wm_row, text="", style="Sub.TLabel")
        self.wm_summary.pack(side="left", padx=10)
        ttk.Button(wm_row, text="自动更新设置", command=self._save_settings).pack(side="left", padx=10)
        self.set_auto = tk.BooleanVar(value=bool(self.settings.get("auto_update", True)))
        ttk.Checkbutton(wm_row, text="启动时检查更新（连接 GitHub）",
                        variable=self.set_auto).pack(side="left")

        info = ttk.LabelFrame(f, text=" 数据存储 ", style="Card.TLabelframe")
        info.pack(fill="x", pady=8)
        self.store_info = ttk.Label(info, text="", style="Sub.TLabel", justify="left")
        self.store_info.pack(anchor="w", padx=8, pady=6)

    def _refresh_settings(self):
        self.set_outdir.set(self.settings.get("outdir", ""))
        self.set_fname.set(self.settings.get("filename_tpl", DEFAULT_FILENAME_TPL))
        self.set_gp_mode.set(bool(self.settings.get("gp_owner_mode")))
        self.set_gp_pw.set(self.settings.get("gp_owner_pw", ""))
        self.set_auto.set(bool(self.settings.get("auto_update", True)))
        self._update_wm_summary()
        self.store_info.configure(
            text="位置：%s\n结构：settings.json · events/ · scenarios/ · runs/\n"
                 "旧版数据：%s"
                 % (self.store.data_dir,
                    "已迁移，保留为 lp_data.json.bak" if self.store.legacy_file.with_name(
                        "lp_data.json.bak").exists() else "无"))

    def _set_pick_outdir(self):
        d = filedialog.askdirectory(title="默认输出目录")
        if d:
            self.set_outdir.set(d)
            self._save_settings()

    def _save_settings(self):
        self.settings["outdir"] = self.set_outdir.get().strip()
        self.settings["filename_tpl"] = self.set_fname.get().strip() or DEFAULT_FILENAME_TPL
        self.settings["gp_owner_mode"] = bool(self.set_gp_mode.get())
        self.settings["gp_owner_pw"] = self.set_gp_pw.get().strip()
        self.settings["auto_update"] = bool(self.set_auto.get())
        ok, err = self.store.save_settings(self.settings)
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        self._update_wm_summary()

    def _persist_settings(self):
        """只把 self.settings 落盘（不依赖「设置」页的 tk 变量）。

        供流水线参数区等非设置页的开关使用（v3.4.4：统一密码勾选原来调
        `_save_settings()`，在非设置页会因 set_outdir 等变量不存在而崩）。
        """
        ok, err = self.store.save_settings(self.settings)
        if not ok:
            messagebox.showerror("保存失败", err)

    # ================= 帮助 =================

    def _build_help(self, parent=None):
        f = parent if parent is not None else self._screens["help"]
        ttk.Label(f, text="帮助", style="Head.TLabel").pack(anchor="w", pady=(0, 6))
        box = ttk.LabelFrame(f, text=" Word 占位符写法 ", style="Card.TLabelframe")
        box.pack(fill="x", pady=4)
        ttk.Label(box, justify="left", text=(
            "在 Word 模板需要替换的位置写占位符，两种格式都支持：\n"
            "    {{LP_Name}}    或    ?LP_Name?\n\n"
            "替换通过 Word 查找替换完成，字体、字号、颜色、表格、页眉页脚、分页全部保持原样。\n"
            "在「文档流水线 → 字段映射」里，每个占位符单独选一个来源：\n"
            "    · 收件人名称 / 水印名称 / 打开口令 / 日期 —— 内置\n"
            "    · 收件人字段 —— 收件人表里的自定义列（可临时加列）\n"
            "    · 事件常量 —— 事件选项里定义的固定值\n"
            "    · 外部数据表 —— 从 Excel 按收件人匹配进来的列\n"
            "表格里有一列「预览值」，直接显示以第一位收件人为例填进去会是什么。\n"
            "注意：权限口令不是可用字段，架构上无法写进邮件或文档。")
        ).pack(anchor="w", padx=10, pady=8)
        flow = ttk.LabelFrame(f, text=" 怎么用 ", style="Card.TLabelframe")
        flow.pack(fill="x", pady=4)
        ttk.Label(flow, justify="left", text=(
            "一次分发 = 一个「事件」，整个流程就三段：\n"
            "    ① 收件人 —— 谁收、多少人、发几封（发送模式）\n"
            "    ② 文档流水线 —— 一张表，每一行 = 一个要附上的东西\n"
            "         · Word-MailMerge：用 Word 模板给每位收件人生成内容不同的文档\n"
            "         · 附加文件：把磁盘上已有的文件放进附件包\n"
            "         每行可勾选「转 PDF / 加水印 / 加密」，并单独设「产出颗粒度」\n"
            "    ③ 邮件模板 —— 主题与正文，{{字段}} 按每位收件人替换\n\n"
            "流水线页是左右分栏：\n"
            "    左栏 = 条目表 + 选中条目的参数区\n"
            "    右栏 = 📎 统一附件（可折叠）+「▶ 处理」按钮 + 收件人↔文件对应表\n"
            "    「▶ 处理」按当前参数全量重跑，生成文件并回填每位收件人的产物路径\n\n"
            "每个人的附件从哪来（三处，作用不同）：\n"
            "    · 流水线的 Word-MailMerge 条目 —— 每人一份、内容各异的文档\n"
            "    · 流水线的「附加文件」条目 —— 手填路径清单，或从 Excel 导入路径（与收件人一一对应）\n"
            "    · 统一附件 / 「附件」页 —— 所有人或指定的人收到的同一份静态文件\n\n"
            "两个「模式」别搞混：\n"
            "    发送模式（事件级）→ 决定几封邮件：每人一份 / 按组合并 / 全部合并\n"
            "    产出颗粒度（条目级）→ 决定这一条产出几份文件：每人一份 / 每组合并一份 / 全员共用一份\n\n"
            "操作顺序：\n"
            "1. 事件库 → 新建事件（或从「场景模板」复制一份配置）\n"
            "2. 收件人：先「下载导入模板」发给业务方填，填好导回来；\n"
            "    导入会先做数据校验，逐行列出问题，确认后才写入\n"
            "3. 文档流水线：点「＋ Word-MailMerge」或「＋ 附加文件」加行；\n"
            "    单击行在左侧改参数；Word-MailMerge 参数区有「说明」「导出要素模板」\n"
            "    「下载示例 Word 模板」；按住行上下拖动改顺序，Ctrl/Shift 多选批量删除\n"
            "4. 附件：不想走流水线时，用它快捷挂载静态文件（全部 / 指定收件人 / 按组）\n"
            "5. 邮件模板：写主题与正文\n"
            "6. 总览：核对每位收件人最终收到什么、邮件长什么样\n"
            "7. 预览与生成：预检 → 生成全部 Outlook 草稿（只建草稿，不发送）\n\n"
            "在事件库里选中任一事件，右侧会画出它的流程图。\n"
            "导出备份：事件库里可导出收件人 Excel、事件 JSON、整库备份。\n\n"
            "生成完成后到 Outlook 草稿箱逐封复核，人工发送。")
        ).pack(anchor="w", padx=10, pady=8)

    # ================= 事件工作台 =================

    def open_event(self, event_id):
        ev, err = self.store.load_event(event_id)
        if err:
            messagebox.showerror("读取失败", err)
            return
        if not ev:
            messagebox.showerror("读取失败", "事件不存在")
            return
        self._wb_event = ev
        self.cur_event_id = event_id
        self._wb_step = 0
        if "event" in self._screens:
            for c in self._screens["event"].winfo_children():
                c.destroy()
            self._build_workbench(self._screens["event"])
        else:
            self.show_screen("event")
        # 关键：重建工作台后必须重新 show_screen 把它 pack 回来。
        # show_screen 会把其它页 pack_forget，若不重新 pack，第二次点
        # 「打开工作台」就是空操作（界面完全没反应）。
        self.show_screen("event")
        self._set_crumb("事件 · %s" % ev.get("name", ""))

    def _build_workbench(self, f):
        # 顶栏：事件名 + 元信息
        top = ttk.Frame(f)
        top.pack(fill="x", pady=(0, 4))
        self.wb_title = ttk.Label(top, text="", style="Head.TLabel")
        self.wb_title.pack(side="left")
        self.wb_meta = ttk.Label(top, text="", style="Sub.TLabel")
        self.wb_meta.pack(side="left", padx=12)
        ttk.Button(top, text="事件信息…", command=self._wb_event_info).pack(side="right")
        ttk.Button(top, text="返回列表", command=lambda: self.show_screen("home")
                   ).pack(side="right", padx=6)

        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)

        # 左侧步骤导航
        nav = ttk.Frame(body, width=168)
        nav.pack(side="left", fill="y")
        nav.pack_propagate(False)
        self._wb_step_buttons = []
        for i, (title, _key) in enumerate(WB_STEPS):
            b = ttk.Button(nav, style="Step.TButton", text="", width=20,
                           command=lambda n=i: self._wb_show_step(n))
            b.pack(fill="x", padx=4, pady=1)
            self._wb_step_buttons.append(b)

        # 右侧内容区
        self._wb_content = ttk.Frame(body)
        self._wb_content.pack(side="left", fill="both", expand=True)

        # 底部操作条
        foot = ttk.Frame(f)
        foot.pack(fill="x", pady=(6, 0))
        self.btn_preflight = ttk.Button(foot, text="预检", command=self._wb_preflight)
        self.btn_preflight.pack(side="left")
        self.btn_generate = ttk.Button(foot, text="生成全部 Outlook 草稿",
                                       command=self._wb_generate)
        self.btn_generate.pack(side="left", padx=6)
        self.btn_rerun = ttk.Button(foot, text="只重跑失败项", command=self._wb_rerun_failed)
        self.btn_rerun.pack(side="left")
        self.wb_progress = ttk.Progressbar(foot, length=180, mode="determinate")
        self.wb_progress.pack(side="right")
        self.wb_status = ttk.Label(foot, text="", style="Sub.TLabel")
        self.wb_status.pack(side="right", padx=10)
        ttk.Label(foot, foreground="#b00",
                  text="草稿只写入 Outlook 草稿箱，不会发送；发送前请人工复核"
                  ).pack(side="left", padx=14)

        self._wb_refresh_nav()
        self._wb_show_step(0)

    def _wb_refresh_nav(self):
        ev = self._wb_event or {}
        mode = lp_v3.send_mode_of(ev)
        mode_label = lp_v3.SEND_MODE_LABELS.get(mode, mode)
        subs = ["%d 位 · %s" % (len(ev.get("recipients", [])), mode_label),
                "%d 条" % len(ev.get("pipeline", [])),
                "%d 个" % len(ev.get("attachments", [])),
                "已配置" if ev.get("mail", {}).get("subject") else "未配置",
                "%d 封" % len(lp_v3.draft_units(ev)),
                "未执行"]
        for i, (title, _k) in enumerate(WB_STEPS):
            self._wb_step_buttons[i].configure(
                text="(%d) %s\n      %s" % (i + 1, title, subs[i]))
        self.wb_title.configure(text=ev.get("name", "事件"))
        self.wb_meta.configure(text="发送模式：%s ｜ 收件人 %d ｜ 条目 %d"
                               % (mode_label,
                                  len(ev.get("recipients", [])),
                                  len(ev.get("pipeline", []))))

    def _wb_show_step(self, idx):
        self._wb_step = idx
        for c in self._wb_content.winfo_children():
            c.destroy()
        builder = [self._wb_step_recipients, self._wb_step_pipeline,
                   self._wb_step_attachments, self._wb_step_mail,
                   self._wb_step_overview, self._wb_step_run][idx]
        builder(self._wb_content)
        self._set_crumb("事件 · %s · %s" % (self._wb_event.get("name", ""),
                                            WB_STEPS[idx][0]))

    def _wb_save(self):
        if not self._wb_event:
            return
        ok, err = self.store.save_event(self._wb_event)
        if not ok:
            messagebox.showerror("保存失败", err)
        self._wb_refresh_nav()

    def _wb_event_info(self):
        ev = self._wb_event
        win = tk.Toplevel(self.root)
        win.title("事件信息")
        win.grab_set()
        rows = [("事件名称", "name"), ("业务类型", "biz_type"),
                ("截止日期", "due_date"), ("备注", "notes")]
        entries = {}
        for i, (label, key) in enumerate(rows):
            ttk.Label(win, text=label + "：").grid(row=i, column=0, sticky="e", padx=6, pady=4)
            e = ttk.Entry(win, width=46)
            e.insert(0, ev.get(key, ""))
            e.grid(row=i, column=1, padx=6, pady=4)
            entries[key] = e

        # 发送模式（只读展示 + 可改）
        i = len(rows)
        ttk.Label(win, text="发送模式：").grid(row=i, column=0, sticky="e", padx=6, pady=4)
        mode_var = tk.StringVar(value=lp_v3.SEND_MODE_LABELS.get(
            lp_v3.send_mode_of(ev), lp_v3.send_mode_of(ev)))
        ttk.Combobox(win, textvariable=mode_var, state="readonly", width=14,
                     values=[l for _k, l, _d in lp_v3.SEND_MODES]
                     ).grid(row=i, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(win, text=lp_v3.SEND_MODE_DESC.get(lp_v3.send_mode_of(ev), ""),
                  foreground="#666").grid(row=i + 1, column=1, sticky="w", padx=6)

        def ok():
            for k, e in entries.items():
                ev[k] = e.get().strip()
            for k, l, _d in lp_v3.SEND_MODES:
                if l == mode_var.get():
                    opts = ev.setdefault("options", {})
                    opts["send_mode"] = k
                    opts["merge_groups"] = (k != "per_lp")
                    break
            self._wb_save()
            win.destroy()

        btns = ttk.Frame(win)
        btns.grid(row=i + 2, column=0, columnspan=2, pady=8)
        ttk.Button(btns, text="保存", command=ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=8)

    # ---------- 步骤 1：收件人 ----------

    def _wb_field_columns(self):
        """收件人自定义字段列（按首次出现顺序）。"""
        cols = []
        for r in (self._wb_event or {}).get("recipients", []):
            for k in (r.get("fields") or {}):
                if k not in cols:
                    cols.append(k)
        return cols

    def _wb_step_recipients(self, parent):
        ev = self._wb_event
        ttk.Label(parent, text="收件人", style="Head.TLabel").pack(anchor="w")

        # --- 发送模式（事件级开关：决定几封邮件 / 几份文档）---
        mode_row = ttk.Frame(parent)
        mode_row.pack(fill="x", pady=(4, 2))
        ttk.Label(mode_row, text="发送模式：").pack(side="left")
        cur_mode = lp_v3.send_mode_of(ev)
        self._wb_mode_var = tk.StringVar(value=lp_v3.SEND_MODE_LABELS.get(cur_mode, cur_mode))
        cb = ttk.Combobox(mode_row, textvariable=self._wb_mode_var, state="readonly",
                          width=14, values=[l for _k, l, _d in lp_v3.SEND_MODES])
        cb.pack(side="left", padx=4)
        cb.bind("<<ComboboxSelected>>", self._wb_change_send_mode)
        self._wb_mode_hint = ttk.Label(mode_row, text="", style="Sub.TLabel")
        self._wb_mode_hint.pack(side="left", padx=8)
        self._wb_update_mode_hint()

        fcols = self._wb_field_columns()
        ttk.Label(parent, text="自定义字段列供文档映射使用%s。"
                  % ("（当前：" + "、".join(fcols) + "）" if fcols else "（还没有）"),
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 6))

        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 4))
        for text, cmd, tip in [
            ("导入 Excel/CSV", self._wb_import_recipients,
             "从 xlsx/csv 批量导入收件人清单。\n"
             "流程：选文件 → 逐行数据校验 → 可视化导入结果 → 写入。\n"
             "按表头名识别列；未识别的列会作为「自定义字段」收进来，供文档映射使用。\n"
             "同名 LP 覆盖更新、异名新增；口令列留空自动生成随机口令。"),
            ("下载导入模板", self._wb_export_template,
             "导出一份带表头、示例行和「填写说明」页的 Excel，\n"
             "发给业务方按说明填写后再导回来。"),
            ("导出收件人 Excel", self._wb_export_recipients,
             "把当前事件的收件人清单导出成 Excel（含自定义字段列与口令）。"),
            ("新增", self._wb_edit_recipient, "手工添加一位收件人。"),
            ("编辑", self._wb_edit_recipient, "编辑选中的收件人（也可直接双击行）。"),
            ("删除", self._wb_delete_recipients, "删除选中的收件人（可多选）。"),
            ("添加自定义字段列", self._wb_add_field_column,
             "给收件人表加一个自定义字段列（Call_Amount / Unfunded 这类），\n"
             "之后在编辑对话框里给每位收件人填值。"),
            ("套用字段模板", self._wb_apply_field_template,
             "用设置里预定义的字段模板（如「催缴场景字段」）一次性铺开多列。"),
            ("检查附件", self._wb_check_attachments, "检查收件人专属附件文件是否存在。"),
            ("补齐缺失口令", self._wb_fill_passwords,
             "为打开口令/权限口令为空的收件人批量生成随机口令。"),
            ("从收件人库添加…", self._wb_add_from_pool,
             "从收件人库（常用收件人）里挑人加进本事件。\n"
             "按名称合并：本事件已有的口令不会被覆盖。"),
            ("存入收件人库", self._wb_save_to_pool,
             "把本事件的收件人整个存进收件人库，方便下个事件复用。"),
        ]:
            b = ttk.Button(bar, text=text, command=cmd)
            b.pack(side="left", padx=2)
            ToolTip(b, tip)

        cols = ["name", "to", "cc", "bcc", "wm", "group"] + fcols + \
            ["open_pw", "owner_pw"]
        heads = ["收件人名称", "To", "CC", "BCC", "水印", "组"] + fcols + \
                ["打开口令", "权限口令(勿发)"]
        # 注意：Treeview 必须建在 frame 里。之前建在 parent 里、却 pack 进 frame，
        # 结果它被排到了 frame 下面，表格被挤成右下角一小块。
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended")
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
            tree.column(c, width=150 if c in ("to", "cc", "bcc") else 100)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.bind("<Double-1>", lambda _e: self._wb_edit_recipient())
        self._wb_rec_tree = tree

        for r in ev.get("recipients", []):
            # v3.4.4：收件人页不再露出附件路径列，values 必须与 columns 一一对应。
            # （v3.4.0~3.4.3 这里多塞了一个 att_txt，Tk 会丢弃最后一个值
            #  → 权限口令的 •••• 被丢掉 + 整行错位 + 明文口令外泄。）
            vals = [r.get("name", ""), r.get("to", ""), r.get("cc", ""),
                    r.get("bcc", ""), r.get("wm", ""), r.get("group", "")]
            vals += [(r.get("fields") or {}).get(k, "") for k in fcols]
            vals += [r.get("open_pw", ""),
                     "••••" if r.get("owner_pw") else ""]
            tree.insert("", "end", values=vals)

        ttk.Label(parent, foreground="#b00", text=(
            "安全提示：本程序同目录的 lp_data/ 含全部权限口令，属敏感文件，"
            "请限定在本机/受控共享目录使用。")).pack(fill="x", pady=(4, 0))

    # ---------- 发送模式 ----------

    def _wb_update_mode_hint(self):
        if not hasattr(self, "_wb_mode_var"):
            return
        label = self._wb_mode_var.get()
        desc = ""
        for _k, l, d in lp_v3.SEND_MODES:
            if l == label:
                desc = d
                break
        n = len((self._wb_event or {}).get("recipients") or [])
        self._wb_mode_hint.configure(text="%s（当前 %d 位收件人）" % (desc, n))

    def _wb_change_send_mode(self, _evt=None):
        """切换发送模式（事件级开关：决定几封邮件 / 几份文档）。"""
        if not self._wb_event:
            return
        label = self._wb_mode_var.get()
        mode = None
        for k, l, _d in lp_v3.SEND_MODES:
            if l == label:
                mode = k
                break
        if not mode:
            return
        opts = self._wb_event.setdefault("options", {})
        if opts.get("send_mode") == mode:
            return
        opts["send_mode"] = mode
        opts["merge_groups"] = (mode != "per_lp")
        self._wb_save()
        self._wb_update_mode_hint()
        n = len(self._wb_event.get("recipients") or [])
        messagebox.showinfo(
            "发送模式已切换",
            "已切换为「%s」。\n\n%s\n\n当前 %d 位收件人 → 将生成 %d 封草稿。"
            % (label, lp_v3.SEND_MODE_DESC.get(mode, ""), n,
               len(lp_v3.draft_units(self._wb_event))))

    def _wb_export_recipients(self):
        """导出当前事件的收件人清单（Excel）。"""
        ev = self._wb_event
        if not ev:
            return
        recips = ev.get("recipients") or []
        if not recips:
            messagebox.showinfo("提示", "这个事件还没有收件人")
            return
        try:
            rows = lp_v3.recipients_to_rows(recips)
        except Exception as e:
            # 之前这里没有兜底：异常被 tkinter 吞掉，exe 是 --noconsole 构建的，
            # 用户点「导出」后什么都不会发生（连报错都看不到）。
            messagebox.showerror("导出失败", "导出收件人清单时出错：\n%s" % e)
            return
        self._write_xlsx_sheets(
            "收件人清单_%s.xlsx" % sanitize_filename(ev.get("name", "事件"), 40),
            [("收件人清单", rows, [36, 20, 28, 16, 16, 22, 34, 14, 14, 24])],
            "已导出",
            "收件人清单已导出。\n\n注意：文件里含打开口令与权限口令，属敏感文件，"
            "请按与 lp_data/ 相同的管控要求保存。")

    # ---------- 导入：模板下载 / 数据校验 / 可视化结果 ----------

    def _wb_import_recipients(self):
        """收件人步骤里的「导入 Excel/CSV」。"""
        if not self._wb_event:
            return
        stats = self._do_import_recipients(self._wb_event)
        if stats:
            self._wb_show_step(self._wb_step)

    def _do_import_recipients(self, ev):
        """导入收件人清单（收件人步骤与事件库共用同一套流程）。

        选文件 → 逐行数据校验 → 可视化导入结果 → 写入事件。
        返回统计 dict；用户取消则返回 None。
        """
        path = filedialog.askopenfilename(title="选择收件人清单",
                                          filetypes=[("表格", "*.xlsx *.csv"),
                                                     ("所有文件", "*.*")])
        if not path:
            return None
        try:
            rows = read_spreadsheet(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return None
        existing = [r.get("name", "") for r in ev.get("recipients", [])]
        result = lp_v3.validate_recipient_rows(rows, existing)
        action = self._show_import_result(result, path)
        if not action:
            return None
        stats = self._apply_import_rows(result, ev,
                                        only_ok=(action == "ok_only"))
        ok, err = self.store.save_event(ev)
        if not ok:
            messagebox.showerror("保存失败", err)
            return None
        stats["field_cols"] = result.get("field_cols") or []
        stats["unknown_cols"] = result.get("unknown_cols") or []
        return stats

    def _apply_import_rows(self, result, ev, only_ok=True):
        """把校验通过的行写进事件。返回统计。"""
        by_name = {r.get("name", ""): r for r in ev.get("recipients", [])}
        updated = added = skipped = 0
        for item in result.get("items") or []:
            rec = item.get("record") or {}
            name = rec.get("name", "")
            if only_ok and item.get("status") == "error":
                skipped += 1
                continue
            if not name:
                skipped += 1
                continue
            cur = by_name.get(name)
            if cur is None:
                ev.setdefault("recipients", []).append(rec)
                by_name[name] = rec
                added += 1
                continue
            # 合并而非替换：导入值为空则保留原值（防口令被清空重发）
            for k, v in rec.items():
                if k == "fields":
                    for fk, fv in (v or {}).items():
                        if fv:
                            cur.setdefault("fields", {})[fk] = fv
                elif k == "attachments":
                    if v:
                        merged = list(cur.get("attachments") or [])
                        for p in v:
                            if p not in merged:
                                merged.append(p)
                        cur["attachments"] = merged
                elif v and v != cur.get(k, ""):
                    cur[k] = v
            updated += 1
        filled = 0
        for r in ev.get("recipients", []):
            if not r.get("open_pw"):
                r["open_pw"] = gen_pw()
                filled += 1
            if not r.get("owner_pw"):
                r["owner_pw"] = gen_pw()
                filled += 1
        return {"updated": updated, "added": added, "skipped": skipped,
                "filled": filled}

    def _show_import_result(self, result, path):
        """可视化导入结果：逐行状态 + 问题清单 + 导出明细。

        返回 'ok_only'（只导无错行）/ 'all'（全导）/ None（取消）。
        """
        win = tk.Toplevel(self.root)
        win.title("导入结果预览")
        win.geometry("1020x580")
        win.minsize(760, 400)
        win.transient(self.root)
        decision = {"action": None}

        items = result.get("items") or []
        n_err = sum(1 for i in items if i.get("status") == "error")
        n_warn = sum(1 for i in items if i.get("status") == "warning")
        n_ok = sum(1 for i in items if i.get("status") == "ok")

        head = ttk.Frame(win)
        head.pack(fill="x", padx=12, pady=(10, 2))
        ttk.Label(head, text="导入结果预览", style="Head.TLabel").pack(side="left")
        ttk.Label(head, text=Path(path).name, style="Sub.TLabel").pack(side="left", padx=10)

        summary = ttk.Frame(win)
        summary.pack(fill="x", padx=12)
        ttk.Label(summary, text="共 %d 行" % len(items)).pack(side="left")
        ttk.Label(summary, text="将导入 %d 条（通过+警告）" % (n_ok + n_warn),
                  foreground="#2f7d32").pack(side="left", padx=14)
        ttk.Label(summary, text="通过 %d" % n_ok, foreground="#2f7d32").pack(side="left")
        ttk.Label(summary, text="警告 %d" % n_warn, foreground="#b8860b").pack(side="left")
        ttk.Label(summary, text="错误 %d" % n_err, foreground="#b00").pack(side="left", padx=14)
        for m in result.get("fatal") or []:
            ttk.Label(summary, text="✕ " + m, foreground="#b00",
                      wraplength=860, justify="left").pack(anchor="w")
        for m in result.get("notes") or []:
            ttk.Label(summary, text="· " + m, foreground="#666",
                      wraplength=860, justify="left").pack(anchor="w")
        if result.get("unknown_cols"):
            ttk.Label(summary, style="Sub.TLabel", wraplength=860, justify="left",
                      text="未识别的列（会作为自定义字段收进收件人）：%s"
                           % "、".join(result["unknown_cols"])).pack(anchor="w", pady=(4, 0))

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=8)
        cols = ["no", "name", "status", "msg"]
        heads = ["行号", "收件人名称", "结果", "问题"]
        widths = {"no": 58, "name": 190, "status": 64, "msg": 520}
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
            tree.column(c, width=widths[c], stretch=(c == "msg"), anchor="w")
        tree.tag_configure("error", foreground="#b00")
        tree.tag_configure("warning", foreground="#b8860b")
        tree.tag_configure("ok", foreground="#2f7d32")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        label_of = {"ok": "通过", "warning": "警告", "error": "错误"}
        for it in items:
            st = it.get("status", "ok")
            msgs = list(it.get("errors") or []) + list(it.get("warnings") or [])
            tree.insert("", "end", values=(it.get("no", ""), it.get("name", ""),
                                          label_of.get(st, st), "；".join(msgs)),
                        tags=(st,))

        foot = ttk.Frame(win)
        foot.pack(fill="x", padx=12, pady=(0, 10))

        def _finish(act):
            decision["action"] = act
            win.destroy()

        ttk.Button(foot, text="导入无错误的 %d 行" % (n_ok + n_warn),
                   command=lambda: _finish("ok_only")).pack(side="left")
        ttk.Button(foot, text="全部导入（含错误行 %d）" % n_err,
                   command=lambda: _finish("all")).pack(side="left", padx=6)
        ttk.Button(foot, text="导出问题明细…",
                   command=lambda: self._export_import_problems(result)
                   ).pack(side="left")
        ttk.Button(foot, text="取消", command=lambda: _finish(None)).pack(side="right")
        win.protocol("WM_DELETE_WINDOW", lambda: _finish(None))
        win.wait_window()
        return decision["action"]

    def _export_import_problems(self, result):
        """把导入问题明细导出成 Excel，方便业务方改完再导。"""
        rows = [["行号", "收件人名称", "级别", "问题"]]
        for it in result.get("items") or []:
            level = {"ok": "通过", "warning": "警告", "error": "错误"}.get(
                it.get("status", ""), it.get("status", ""))
            msgs = list(it.get("errors") or []) + list(it.get("warnings") or [])
            if not msgs:
                continue
            rows.append([it.get("no", ""), it.get("name", ""), level, "；".join(msgs)])
        if len(rows) == 1:
            rows.append(["—", "—", "—", "没有发现问题"])
        self._write_xlsx_sheets(
            "导入问题明细.xlsx",
            [("问题明细", rows, [10, 30, 10, 90])],
            "已导出", "问题明细已导出，按提示修改后重新导入即可。")

    # ---------- 导出 / 备份 ----------

    def _lib_import_recipients(self):
        """事件库里的「导入收件人」：导入到当前选中的事件。"""
        eid = self._selected_event_id()
        if not eid:
            messagebox.showinfo("提示", "请先在事件表里选中一个事件")
            return
        ev, err = self.store.load_event(eid)
        if not ev:
            messagebox.showerror("读取失败", err or "事件不存在")
            return
        stats = self._do_import_recipients(ev)
        if not stats:
            return
        msg = "导入完成：新增 %d，更新 %d" % (stats["added"], stats["updated"])
        if stats["skipped"]:
            msg += "，跳过 %d 行错误" % stats["skipped"]
        if stats["filled"]:
            msg += "。自动生成缺失口令 %d 个" % stats["filled"]
        if stats["field_cols"]:
            msg += "\n收进自定义字段列：" + "、".join(stats["field_cols"])
        messagebox.showinfo("导入完成", msg)
        self._refresh_home()

    def _lib_export_recipients(self):
        """导出选中事件的收件人清单（Excel）。"""
        eid = self._selected_event_id()
        if not eid:
            messagebox.showinfo("提示", "请先在事件表里选中一个事件")
            return
        ev, err = self.store.load_event(eid)
        if not ev:
            messagebox.showerror("读取失败", err or "事件不存在")
            return
        recips = ev.get("recipients") or []
        if not recips:
            messagebox.showinfo("提示", "这个事件还没有收件人")
            return
        rows = lp_v3.recipients_to_rows(recips)
        name = "收件人清单_%s.xlsx" % sanitize_filename(ev.get("name", "事件"), 40)
        self._write_xlsx_sheets(
            name, [("收件人清单", rows, [36, 20, 28, 16, 16, 22, 34, 14, 14, 24])],
            "已导出",
            "收件人清单已导出。\n\n注意：文件里含打开口令与权限口令，属于敏感文件，"
            "请按与 lp_data/ 相同的管控要求保存。")

    def _lib_export_all_recipients(self):
        """导出全部事件的收件人（合并，带「所属事件」列）。"""
        events = self.store.list_events()
        if not events:
            messagebox.showinfo("提示", "还没有事件")
            return
        rows = [["所属事件"] + list(lp_v3.EXPORT_BASE_HEADERS)]
        for e in events:
            for r in (e.get("recipients") or []):
                one = lp_v3.recipients_to_rows([r])
                rows.append([e.get("name", "")] + one[1])
        self._write_xlsx_sheets(
            "全部收件人清单.xlsx",
            [("全部收件人", rows, [26, 36, 20, 28, 16, 16, 22, 34, 14, 14, 24])],
            "已导出", "全部事件的收件人已导出（含口令，属敏感文件）。")

    def _lib_export_event(self):
        """导出单个事件配置（JSON 备份，可再导入还原）。"""
        eid = self._selected_event_id()
        if not eid:
            messagebox.showinfo("提示", "请先在事件表里选中一个事件")
            return
        ev, err = self.store.load_event(eid)
        if not ev:
            messagebox.showerror("读取失败", err or "事件不存在")
            return
        payload = lp_v3.event_export_payload(ev)
        path = filedialog.asksaveasfilename(
            title="导出事件配置", defaultextension=".json",
            initialfile="事件_%s.json" % sanitize_filename(ev.get("name", "事件"), 40),
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            return
        messagebox.showinfo("已导出",
                            "事件配置已导出。\n\n还原方式：事件库 → 导入收件人… 旁边的"
                            "「导入事件备份…」（或直接把这个 JSON 放回 lp_data/events/）。")

    def _lib_export_backup(self):
        """导出整库备份（设置 + 事件 + 场景 + 执行记录）。"""
        data = lp_v3.backup_payload(self.store)
        path = filedialog.asksaveasfilename(
            title="备份全部数据", defaultextension=".json",
            initialfile="doc-mailer_备份_%s.json" % datetime.now().strftime("%Y%m%d-%H%M%S"),
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            return
        messagebox.showinfo(
            "已备份",
            "已备份 %d 个事件、%d 个场景、%d 条执行记录。\n\n"
            "注意：备份含全部权限口令，属敏感文件。"
            % (len(data["events"]), len(data["scenarios"]), len(data["runs"])))

    def _lib_import_event(self):
        """从 JSON 备份还原一个事件（新建，不覆盖同名事件）。"""
        path = filedialog.askopenfilename(title="导入事件备份",
                                          filetypes=[("JSON", "*.json"),
                                                     ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        ok, ev, err = lp_v3.event_import_payload(payload)
        if not ok:
            messagebox.showerror("导入失败", err)
            return
        ev["id"] = ""
        ev["name"] = (ev.get("name") or "导入事件") + "（导入）"
        ev["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ev["updated_at"] = ev["created_at"]
        ev["status"] = ""
        ok, err = self.store.save_event(ev)
        if not ok:
            messagebox.showerror("保存失败", err)
            return
        messagebox.showinfo("导入完成", "事件「%s」已导入。" % ev["name"])
        self._refresh_home()

    # ---------- 模板下载 ----------

    def _wb_export_template(self):
        """下载收件人清单导入模板（v3.4.4：不再含「附件路径」列）。"""
        headers = lp_v3.import_template_headers(["Call_Amount", "Unfunded", "Due_Date"])
        row = ["示例：苏州某某科技有限公司", "（留空=同收件人名称）", "a@x.com;b@x.com",
               "cc@x.com", "bcc@x.com", "（留空=用邮件模板）",
               "（留空=不合并）", "（留空=自动生成）", "（留空=自动生成，勿发给LP）",
               "200,000", "800,000", "2026-11-15"]
        guide = [
            ["列名", "必填", "说明"],
            ["收件人名称", "是", "收件单位全称；识别依据，也是水印默认文字"],
            ["水印名称", "否", "留空 = 同收件人名称"],
            ["收件人", "是", "收件邮箱；多个用 ; 分隔（英文逗号是文件名的一部分，不做分隔）"],
            ["抄送 / 密送", "否", "邮箱；多个用 ; 分隔"],
            ["邮件标题", "否", "留空 = 用邮件模板里的主题"],
            ["打开口令", "否", "留空 = 自动生成随机口令（收件人打开文件用）"],
            ["权限口令", "否", "留空 = 自动生成（控制打印权限，绝不会进入邮件）"],
            ["合并组", "否", "仅「按组合并」发送模式生效；同组共用一封邮件、附件包合并"],
            ["其它任意列", "否", "作为自定义字段收进收件人，可在流水线的字段映射里用到"],
        ]
        widths = [36, 20, 28, 16, 16, 22, 14, 14, 24, 14, 14, 14]
        self._write_xlsx_sheets(
            "收件人清单导入模板.xlsx",
            [("收件人清单", [headers, row], widths),
             ("填写说明", guide, [18, 8, 78])],
            "模板已导出",
            "把模板发给业务方填写，填完再导回来。\n"
            "表头列名即识别依据；未识别的列会作为自定义字段收进来，"
            "可在流水线里映射到 Word 占位符。\n"
            "口令列留空会自动生成随机口令。")

    def _write_xlsx(self, default_name, rows, ok_title, ok_msg):
        """单表 Excel 导出（兼容旧调用）。"""
        return self._write_xlsx_sheets(
            default_name, [("Sheet1", rows, [36, 20, 28, 16, 16, 22, 34, 14, 24, 14, 14, 14])],
            ok_title, ok_msg)

    def _write_xlsx_sheets(self, default_name, sheets, ok_title, ok_msg):
        """导出多页 Excel。sheets = [(表名, 二维列表, 列宽列表), ...]"""
        path = filedialog.asksaveasfilename(title="保存", defaultextension=".xlsx",
                                            initialfile=default_name,
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return None
        import openpyxl
        from openpyxl.utils import get_column_letter
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for title, rows, widths in sheets:
            ws = wb.create_sheet(title=str(title)[:31])
            for row in rows or []:
                ws.append(list(row))
            for i, w in enumerate(widths or []):
                ws.column_dimensions[get_column_letter(i + 1)].width = w
            try:
                for c in ws[1]:
                    c.font = openpyxl.styles.Font(bold=True)
            except Exception:
                pass
        try:
            wb.save(path)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return None
        messagebox.showinfo(ok_title, ok_msg)
        return path

    def _wb_selected_recipients(self):
        tree = getattr(self, "_wb_rec_tree", None)
        if tree is None:
            return []
        ev = self._wb_event
        idxs = sorted((tree.index(i) for i in tree.selection()), reverse=True)
        return [ev["recipients"][i] for i in idxs]

    def _recipient_dialog(self, rec, field_cols, title, on_save, no_field_hint=""):
        """收件人编辑对话框 —— 事件里与收件人库里**共用同一套表单**。

        rec: 初始值；field_cols: 自定义字段列名；
        on_save(out): 保存回调，返回 True 表示关闭对话框。
        """
        win = tk.Toplevel(self.root)
        win.title(title)
        win.grab_set()
        rows = [("收件人名称（水印默认用同名）", "name"), ("水印名称", "wm"),
                ("收件人To（多人用;分隔）", "to"), ("抄送CC（可空）", "cc"),
                ("密送BCC（可空）", "bcc"),
                ("合并发送组（同名=一封邮件，可空）", "group"),
                ("邮件标题（空=用模板）", "subject"),
                ("打开口令（发LP）", "open_pw"), ("权限口令（勿发LP）", "owner_pw")]
        entries = {}
        for i, (label, key) in enumerate(rows):
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="e", padx=6, pady=4)
            e = ttk.Entry(win, width=52, show="*" if key == "owner_pw" else "")
            e.insert(0, rec.get(key, ""))
            e.grid(row=i, column=1, padx=6, pady=4)
            entries[key] = e
        r0 = len(rows)
        ttk.Label(win, text="自定义字段", foreground="#444").grid(
            row=r0, column=0, sticky="e", padx=6, pady=(8, 2))
        field_entries = {}
        fcols = field_cols or []
        if not fcols:
            ttk.Label(win, text=no_field_hint or "（没有自定义字段列）",
                      foreground="#888").grid(row=r0, column=1, sticky="w", padx=6)
        for j, k in enumerate(fcols):
            ttk.Label(win, text=k).grid(row=r0 + 1 + j, column=0, sticky="e", padx=6, pady=2)
            e = ttk.Entry(win, width=52)
            e.insert(0, (rec.get("fields") or {}).get(k, ""))
            e.grid(row=r0 + 1 + j, column=1, padx=6, pady=2)
            field_entries[k] = e

        def ok():
            out = {k: e.get().strip() for k, e in entries.items()}
            if not out["name"]:
                messagebox.showwarning("必填", "收件人名称不能为空", parent=win)
                return
            if out["owner_pw"] and out["owner_pw"] == out["open_pw"]:
                messagebox.showwarning("口令", "权限口令不能与打开口令相同", parent=win)
                return
            out["wm"] = out["wm"] or out["name"]
            # v3.4.4：附件路径不再在对话框里维护，保存时原样保留旧数据
            # （流水线「路径来源=收件人附件」仍兼容旧数据，不能在这里被清空）。
            out["attachments"] = list(rec.get("attachments") or [])
            out["fields"] = {k: e.get().strip() for k, e in field_entries.items()}
            if on_save(out):
                win.destroy()

        btns = ttk.Frame(win)
        btns.grid(row=r0 + 1 + len(fcols), column=0, columnspan=2, pady=8)
        ttk.Button(btns, text="保存", command=ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=8)
        return win

    def _wb_edit_recipient(self):
        tree = getattr(self, "_wb_rec_tree", None)
        ev = self._wb_event
        sel = tree.selection() if tree else ()
        rec = dict(ev["recipients"][tree.index(sel[0])]) if sel else {
            "name": "", "wm": "", "to": "", "cc": "", "bcc": "",
            "open_pw": gen_pw(), "owner_pw": gen_pw(), "fields": {},
            "attachments": [], "group": ""}

        def save(out):
            if sel:
                ev["recipients"][tree.index(sel[0])] = out
            else:
                ev["recipients"].append(out)
            self._wb_save()
            self._wb_show_step(self._wb_step)
            return True

        self._recipient_dialog(
            rec, self._wb_field_columns(),
            "编辑收件人" if sel else "新增收件人", save,
            "（还没有自定义字段列，可先在收件人页点「添加自定义字段列」）")

    def _wb_add_from_pool(self):
        """从收件人库挑人加进当前事件。"""
        ev = self._wb_event
        if not ev:
            return
        pool = self.store.list_pool()
        if not pool:
            messagebox.showinfo("收件人库是空的",
                                "收件人库里还没有人。\n"
                                "可以先在「事件库 → 收件人库」标签页里新建，"
                                "或用「存入收件人库」把某个事件的收件人存进去。")
            return
        win = tk.Toplevel(self.root)
        win.title("从收件人库添加")
        win.geometry("620x520")
        win.grab_set()
        ttk.Label(win, text=(
            "勾选要加进「%s」的收件人。\n"
            "按名称合并：本事件里已有的收件人口令不会被覆盖，只补空缺。"
            % ev.get("name", "")), justify="left").pack(anchor="w", padx=12, pady=(10, 4))
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12)
        tree = ttk.Treeview(frame, columns=["name", "to"], show="headings",
                            selectmode="extended", height=16)
        tree.heading("name", text="收件人名称")
        tree.heading("to", text="To")
        tree.column("name", width=280)
        tree.column("to", width=260)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        ids = {}
        for i, r in enumerate(pool):
            iid = "p%d" % i
            ids[iid] = r
            tree.insert("", "end", iid=iid, values=(r.get("name", ""), r.get("to", "")))

        def ok():
            recs = [ids[i] for i in tree.selection() if i in ids]
            if not recs:
                messagebox.showwarning("未选择", "请先勾选要添加的收件人", parent=win)
                return
            new_ev, added, updated = lp_v3.merge_recipients_into_event(ev, recs)
            ok2, err = self.store.save_event(new_ev)
            if not ok2:
                messagebox.showerror("保存失败", err, parent=win)
                return
            ev["recipients"] = new_ev["recipients"]
            self._wb_save()
            self._wb_show_step(self._wb_step)
            messagebox.showinfo("已添加",
                                "新增 %d 位，按名称合并 %d 位（已有口令保留）。"
                                % (added, updated), parent=win)
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(pady=10)
        ttk.Button(btns, text="添加", command=ok).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=6)

    def _wb_save_to_pool(self):
        """把当前事件的收件人整个存进收件人库。"""
        ev = self._wb_event
        recips = ev.get("recipients") or []
        if not recips:
            messagebox.showinfo("提示", "这个事件还没有收件人")
            return
        pool = self.store.list_pool()
        added = updated = 0
        for rec in recips:
            src = dict(rec)
            src.pop("id", None)
            pool, new_rec, is_update = lp_v3.pool_merge_recipient(pool, src)
            if new_rec is not None:
                ok, err, _saved = self.store.save_pool_recipient(new_rec)
                if not ok:
                    messagebox.showerror("保存失败", err)
                    return
                added += 1
            elif is_update:
                updated += 1
        self._refresh_pool_tab()
        messagebox.showinfo("已存入收件人库",
                            "新增 %d 位，按名称合并 %d 位（库里已有的口令保留）。"
                            % (added, updated))

    def _wb_delete_recipients(self):
        tree = getattr(self, "_wb_rec_tree", None)
        ev = self._wb_event
        sel = tree.selection() if tree else ()
        if not sel:
            return
        idxs = sorted((tree.index(i) for i in sel), reverse=True)
        names = [ev["recipients"][i]["name"] for i in reversed(idxs)]
        preview = "\n".join(names[:6]) + ("…" if len(names) > 6 else "")
        if messagebox.askyesno("删除", "确认删除选中的 %d 条？\n\n%s" % (len(names), preview)):
            for i in idxs:
                del ev["recipients"][i]
            self._wb_save()
            self._wb_show_step(self._wb_step)

    def _wb_add_field_column(self):
        name = simpledialog.askstring("添加自定义字段列",
                                       "字段名（如 Call_Amount）：")
        if not name:
            return
        name = name.strip()
        ev = self._wb_event
        if not any(name in (r.get("fields") or {}) for r in ev["recipients"]):
            for r in ev["recipients"]:
                r.setdefault("fields", {}).setdefault(name, "")
        self._wb_save()
        self._wb_show_step(self._wb_step)

    def _wb_apply_field_template(self):
        templates = self.settings.get("field_templates") or []
        if not templates:
            messagebox.showinfo("字段模板", "还没有预定义的字段模板")
            return
        names = [t.get("name", "") for t in templates]
        pick = simpledialog.askstring("套用字段模板",
                                       "选择字段模板：\n\n" + "\n".join(names),
                                       initialvalue=names[0])
        if not pick:
            return
        tpl = next((t for t in templates if t.get("name") == pick), None)
        if tpl is None:
            return
        ev = self._wb_event
        cols = tpl.get("columns") or []
        for r in ev["recipients"]:
            for c in cols:
                r.setdefault("fields", {}).setdefault(c, "")
        self._wb_save()
        self._wb_show_step(self._wb_step)
        messagebox.showinfo("已套用", "已添加字段列：" + "、".join(cols))

    def _fingerprint_notice(self):
        """同封邮件内不同路径但内容相同的附件提示（只提示，不去重、不改路径）。"""
        try:
            return lp_v3.attachment_fingerprint_warnings(self._wb_event)
        except Exception:
            return []

    def _wb_check_attachments(self):
        total, missing = 0, []
        for rec in self._wb_event.get("recipients", []):
            for path in rec.get("attachments") or []:
                total += 1
                if not Path(path).exists():
                    missing.append("%s：%s" % (rec.get("name", ""), path))
        fp = self._fingerprint_notice()
        fp_text = ("\n\n[附件重复提示] 同一封邮件内以下文件内容完全相同"
                   "（不同路径），两份都会保留，请人工确认：\n"
                   + "\n".join(fp)) if fp else ""
        if not missing:
            (messagebox.showwarning if fp else messagebox.showinfo)(
                "附件检查",
                "全部就绪：%d 个附件路径均存在。" % total + fp_text)
        else:
            lines = "\n".join(missing[:15]) + ("…" if len(missing) > 15 else "")
            messagebox.showwarning("附件检查",
                                   "%d/%d 个就绪，以下 %d 个不存在：\n\n%s\n\n"
                                   "缺失的附件在生成草稿时会被跳过并记入日志。"
                                   % (total - len(missing), total, len(missing), lines)
                                   + fp_text)

    def _wb_fill_passwords(self):
        n = 0
        for rec in self._wb_event.get("recipients", []):
            if not rec.get("open_pw"):
                rec["open_pw"] = gen_pw(); n += 1
            if not rec.get("owner_pw"):
                rec["owner_pw"] = gen_pw(); n += 1
        self._wb_save()
        self._wb_show_step(self._wb_step)
        messagebox.showinfo("补齐口令", "新生成 %d 个口令" % n)

    # ---------- 步骤 2：文档流水线 ----------

# ---------- 步骤 2：文档流水线（v3.1：条目表格） ----------

    def _wb_pipeline_items(self):
        return (self._wb_event or {}).setdefault("pipeline", [])

    def _wb_item_by_id(self, iid):
        for it in self._wb_pipeline_items():
            if it.get("id") == iid:
                return it
        return None

    def _wb_selected_item_id(self):
        tree = getattr(self, "_wb_pipe_tree", None)
        if tree is not None:
            sel = tree.selection()
            if sel:
                return sel[0]
        return getattr(self, "_wb_sel_item", None)

    # ---------- 表格渲染 ----------

    def _wb_item_source_text(self, item):
        params = item.get("params") or {}
        if (item.get("kind") or "generate") == "generate":
            tp = (params.get("template_path") or "").strip()
            return Path(tp).name if tp else "（未选模板）"
        # 附加文件：按「路径来源」显示，不是按手填清单显示
        src = str(params.get("path_source") or lp_v3.PATH_SOURCE_MANUAL)
        if src != lp_v3.PATH_SOURCE_MANUAL:
            return "按" + lp_v3.path_source_label(src)
        files = params.get("files") or []
        if not files:
            return "（未选文件）"
        if len(files) == 1:
            return Path(str(files[0])).name
        return "%s 等 %d 个文件" % (Path(str(files[0])).name, len(files))

    def _wb_item_proc_text(self, item):
        params = item.get("params") or {}
        procs = []
        if (item.get("kind") or "generate") == "generate" and params.get("to_pdf"):
            procs.append("PDF")
        if params.get("watermark"):
            procs.append("水印")
        if params.get("encrypt"):
            procs.append("加密")
        return "·".join(procs) if procs else "—"

    def _wb_fill_pipe_table(self, keep_sel=None):
        # 按事件数据重画条目表格（尽量保持当前选中行）。
        tree = getattr(self, "_wb_pipe_tree", None)
        if tree is None:
            return
        sel = keep_sel or getattr(self, "_wb_sel_item", None)
        tree.delete(*tree.get_children())
        items = self._wb_pipeline_items()
        for i, it in enumerate(items):
            iid = it.get("id") or ("it-%d" % i)
            it["id"] = iid
            tree.insert("", "end", iid=iid, values=(
                str(i + 1),
                it.get("name") or ITEM_KIND_LABELS.get(it.get("kind"), "条目"),
                ITEM_KIND_LABELS.get(it.get("kind"), ""),
                self._wb_item_source_text(it),
                self._wb_item_proc_text(it),
                GRANULARITY_LABELS.get(it.get("granularity") or "per_lp", ""),
                "✓" if it.get("enabled", True) else "—",
            ))
        if sel and tree.exists(sel):
            tree.selection_set(sel)
            tree.focus(sel)
        elif items:
            first = items[0].get("id")
            if first and tree.exists(first):
                tree.selection_set(first)
        cnt = getattr(self, "_wb_pipe_count", None)
        if cnt is not None:
            on = len([i for i in items if i.get("enabled", True)])
            cnt.configure(text="共 %d 个条目（启用 %d）" % (len(items), on))

    # ---------- 流水线页 ----------

    def _toggle_uni(self):
        uni = getattr(self, "_uni_frame", None)
        holder = getattr(self, "_uni_btn_holder", None)
        if uni is None or holder is None:
            return
        anchor = getattr(self, "_uni_anchor", None)
        shown = getattr(self, "_uni_shown", False)
        if shown:
            uni.pack_forget()
            if anchor is not None:
                holder.pack(fill="x", pady=(0, 4), before=anchor)
            else:
                holder.pack(fill="x", pady=(0, 4))
            self._uni_shown = False
        else:
            holder.pack_forget()
            if anchor is not None:
                uni.pack(fill="x", pady=(0, 4), before=anchor)
            else:
                uni.pack(fill="x", pady=(0, 4))
            self._uni_shown = True

    def _wb_map_edit(self, event):
        """双击对应表「文件路径」列 → 行内 Entry 编辑。"""
        tree = self._wb_map_tree
        if tree.identify_column(event.x) not in ("#3", "#4"):
            return
        iid = tree.identify_row(event.y)
        if not iid:
            return
        col = tree.identify_column(event.x)
        x, y, w, h = tree.bbox(iid, col)
        cur = tree.set(iid, col)
        e = tk.Entry(tree, width=max(10, w // 8))
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, cur)
        try:
            e.focus_force()     # v3.4.4：部分环境 focus_set 不生效 → Return 不提交
        except Exception:
            e.focus_set()

        def commit(_e=None):
            try:
                tree.set(iid, col, e.get().strip())
            finally:
                try:
                    e.destroy()
                except Exception:
                    pass

        for _seq in ("<Return>", "<KP_Enter>", "<FocusOut>"):
            e.bind(_seq, commit)

    def _wb_map_save_edits(self):
        """把人工改过的路径存进事件 process_overrides。

        单文件存字符串；多文件存路径列表（不要把 " / ".join 的展示串当路径存进去）。
        """
        tree = self._wb_map_tree
        overrides = {}
        for iid in tree.get_children():
            vals = tree.item(iid)["values"]
            if len(vals) < 3:
                continue
            raw = str(vals[2]).strip()
            if not raw or raw == "—":
                continue
            paths = [p.strip() for p in raw.split(" / ") if p.strip() and p.strip() != "—"]
            if not paths:
                continue
            overrides[str(vals[0])] = paths[0] if len(paths) == 1 else paths
        self._wb_event["process_overrides"] = overrides
        self._wb_save()
        messagebox.showinfo("已保存",
                            "自定义路径已保存（%d 条）。\n"
                            "「总览」页生成草稿时将优先使用这些路径。" % len(overrides))

    def _wb_process(self):
        """④⑤「处理」按钮：按上方条目参数全量重跑流水线（不用历史数据），
        完成后把每位收件人的产出文件路径回填到下方对应表。只生成文件，
        不建 Outlook 草稿（草稿在「预览与生成」步骤）。"""
        if self.running:
            messagebox.showinfo("正在执行", "上一次处理还没结束，请等它完成。")
            return
        ev = self._wb_event
        errs, _warns, _info = lp_v3.preflight(ev, self.store)
        if errs:
            messagebox.showerror("预检未通过（未开始处理）",
                                 "以下问题必须先修正：\n\n"
                                 + "\n".join("· " + e for e in errs[:15]))
            return
        outdir = Path((ev.get("options") or {}).get("outdir")
                      or self.settings.get("outdir")
                      or (Path.home() / "Documents" / "文档分发输出"))
        outdir.mkdir(parents=True, exist_ok=True)
        ev.setdefault("options", {})["outdir"] = str(outdir)
        ev["process_results"] = {}   # 全量更新：清空上次记录，避免历史污染
        self._wb_save()

        self.running = True
        self._cancel = threading.Event()
        self.log("===== 处理开始：按当前参数全量生成 =====")

        def job(get_word):
            run = lp_v3.run_event(self.store, ev, create_drafts=False,
                                  outdir=str(outdir),
                                  log=lambda m: self._q.put(("log", m)),
                                  progress=lambda p_: self._q.put(("prog", p_)),
                                  word=get_word())
            self._q.put(("run_done", (run, ev)))

        self._submit_job(job)

    def _apply_process_results(self, run, ev=None):
        """把 Run 里每位收件人的产物文件回填到对应表（每次处理全量刷新）。

        Run items 的字段：name(收件人或组名)/members/files(最终产物路径)。
        按组合并的 Run item 拆回每位成员（各自名下记录组共享的文件）。
        """
        ev = ev or self._wb_event
        if ev is None:
            return
        results = {}
        for it in run.get("items") or []:
            files = [str(f) for f in it.get("files") or []]
            members = it.get("members") or []
            if it.get("kind") == "group" and len(members) > 1:
                for m in members:      # 组共享：每个成员名下都记一份
                    results.setdefault(m, []).extend(files)
            else:
                for m in members or ([it.get("name")] if it.get("name") else []):
                    results.setdefault(m, []).extend(files)
        # 按收件人清单顺序去重
        for r in ev.get("recipients") or []:
            seen, ded = set(), []
            for f in results.get(r.get("name", "")) or []:
                if f.lower() not in seen:
                    seen.add(f.lower()); ded.append(f)
            if r.get("name", "") in results or ded:
                results[r.get("name", "")] = ded
        ev["process_results"] = results
        self._wb_save()
        self.log("===== 处理完成：%d 位收件人的文件路径已更新到对应表 ====="
                 % len(results))
        self._wb_render_mapping_table()

    def _wb_render_mapping_table(self):
        """刷新收件人 ↔ 文件路径对应表。"""
        tree = getattr(self, "_wb_map_tree", None)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        ev = self._wb_event
        results = ev.get("process_results") or {}
        done = bool(ev.get("process_results"))
        for r in ev.get("recipients") or []:
            name = r.get("name", "")
            files = results.get(name) or []
            state = "已处理" if files else ("失败/未处理" if done else "未处理")
            tree.insert("", "end", values=[
                name,
                "\n".join(Path(p).name for p in files) if files else "—",
                " / ".join(str(p) for p in files) if files else "—",
                state])

    def _wb_uni_add_file(self):
        files = filedialog.askopenfilenames(title="选择统一附件（所有人同一份）")
        if not files:
            return
        box = getattr(self, "_wb_uni_text", None)
        if box is None:
            return
        for f in files:
            box.insert("end", str(f) + "\n")
        self._wb_uni_save()

    def _wb_uni_save(self):
        box = getattr(self, "_wb_uni_text", None)
        if box is None:
            return
        atts = []
        for line in box.get("1.0", "end").splitlines():
            pth = line.strip()
            if pth and Path(pth).exists():
                atts.append(lp_v3.normalize_attachment(
                    {"name": Path(pth).name, "path": pth, "scope": "all"}))
            elif pth:
                messagebox.showwarning("附件不存在", "跳过不存在的文件：\n" + pth)
        self._wb_event["attachments"] = atts
        self._wb_save()

    def _wb_step_pipeline(self, parent):
        ttk.Label(parent, text="文档流水线", style="Head.TLabel").pack(anchor="w")
        ttk.Label(parent, text=(
            "每一行 = 一个要附上的东西，只有两种：\n"
            "    ① Word-MailMerge —— 用 Word 模板给每位收件人生成内容各异的文档\n"
            "    ② 附加文件 —— 把磁盘上已有的文件放进附件包\n"
            "单击行 → 左侧参数区；按住行上下拖动 → 改顺序；Ctrl/Shift 多选 → 批量删除。"),
            style="Sub.TLabel", justify="left").pack(anchor="w", pady=(0, 6))

        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 4))
        ttk.Button(bar, text="＋ Word-MailMerge",
                   command=lambda: self._wb_add_item("generate")).pack(side="left")
        ttk.Button(bar, text="＋ 附加文件",
                   command=lambda: self._wb_add_item("attach")).pack(side="left", padx=4)
        ttk.Button(bar, text="删除选中",
                   command=self._wb_del_selected_items).pack(side="left")
        ttk.Button(bar, text="清空", command=self._wb_clear_items).pack(side="left")
        self._wb_pipe_count = ttk.Label(bar, foreground="#888", text="")
        self._wb_pipe_count.pack(side="left", padx=10)

        # ④ 左右分栏：v3.4.4 起所有控件**直接建在对应栏里**。
        # （v3.4.3 用 pack(in_=left_col) 只是把窗口塞进左栏的打包列表，父窗口
        #  仍是 parent，PanedWindow 又后创建、z 序在上，结果左栏内容被整个盖住。）
        pw2 = ttk.PanedWindow(parent, orient="horizontal")
        pw2.pack(fill="both", expand=True, pady=(0, 4))
        # v3.4.5：给两栏显式初始宽度，避免默认窗口下左栏被右栏挤到按钮被裁掉
        left_col = ttk.Frame(pw2, width=540)
        pw2.add(left_col, weight=5)
        right_col = ttk.Frame(pw2, width=400)
        pw2.add(right_col, weight=3)

        # ---------- 左栏：条目表 + 参数区 ----------
        frame = ttk.Frame(left_col)
        frame.pack(fill="both", expand=True, pady=(0, 4))
        cols = ["no", "name", "kind", "source", "proc", "gran", "on"]
        heads = ["#", "条目名称", "类型", "来源 / 模板", "处理", "产出颗粒度", "启用"]
        tree = ttk.Treeview(frame, columns=cols, show="headings",
                            selectmode="extended", height=6)
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
        tree.column("no", width=36, stretch=False, anchor="center")
        tree.column("name", width=130)
        tree.column("kind", width=78, stretch=False)
        tree.column("source", width=150)
        tree.column("proc", width=96, stretch=False)
        tree.column("gran", width=104, stretch=False)
        tree.column("on", width=44, stretch=False, anchor="center")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        # ⑥ v3.4.4：补上横向滚动条（列宽被窗口截断时可以拖）
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self._wb_pipe_tree = tree
        self._wb_pipe_hsb = hsb

        tree.bind("<<TreeviewSelect>>", self._wb_on_pipe_select)
        tree.bind("<Double-1>", self._wb_on_pipe_double)
        tree.bind("<ButtonPress-1>", self._wb_pipe_press)
        tree.bind("<B1-Motion>", self._wb_pipe_motion)
        tree.bind("<ButtonRelease-1>", self._wb_pipe_release)

        self._wb_params_frame = ttk.Frame(left_col)
        self._wb_params_frame.pack(fill="both", expand=True, pady=(0, 4))

        # ---------- 右栏：统一附件（可折叠）+ 处理按钮 + 对应表 ----------
        self._uni_btn_holder = ttk.Frame(right_col)
        self._uni_btn_holder.pack(fill="x", pady=(0, 4))
        ttk.Button(self._uni_btn_holder, text="📎 统一附件（所有人收同一份文件）…",
                   command=self._toggle_uni).pack(side="left")

        proc = ttk.LabelFrame(right_col, text=" 处理与文件对应（收件人 → 将收到的文件） ",
                              style="Card.TLabelframe")
        proc.pack(fill="both", expand=True)
        self._uni_anchor = proc      # 统一附件展开时排在处理区上方

        uni = ttk.LabelFrame(right_col, text=" 统一附件（所有收件人都收到的同一份文件） ",
                             style="Card.TLabelframe")
        self._uni_frame = uni
        self._uni_shown = False
        uni_row = ttk.Frame(uni)
        uni_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(uni_row, text="文件（每行一个，全部收件人共用）：").pack(side="left", anchor="n")
        self._wb_uni_text = tk.Text(uni_row, height=3, width=40, wrap="none")
        self._wb_uni_text.pack(side="left", padx=4)
        for _a in ((self._wb_event or {}).get("attachments") or []):
            self._wb_uni_text.insert("end", str(_a.get("path") or "") + "\n")
        uni_btns = ttk.Frame(uni_row)
        uni_btns.pack(side="left", anchor="n", padx=4)
        ttk.Button(uni_btns, text="添加文件…", width=10,
                   command=self._wb_uni_add_file).pack(side="left")
        ttk.Button(uni_btns, text="保存", width=6,
                   command=self._wb_uni_save).pack(side="left", pady=4)
        ttk.Label(uni, foreground="#888", wraplength=560, justify="left",
                  text="每人专属的内容（如带各自水印/口令的 PDF）用左侧流水线条目处理；"
                       "这里只放所有人同一份的文件，会进入每位收件人的附件包。"
                  ).pack(anchor="w", padx=8, pady=(0, 4))

        prow = ttk.Frame(proc)
        prow.pack(fill="x", padx=8, pady=4)
        self._wb_process_btn = ttk.Button(prow, text="▶ 处理（按上方参数全量生成）",
                                          command=self._wb_process)
        self._wb_process_btn.pack(side="left")
        self._wb_progress2 = ttk.Progressbar(prow, length=110, mode="determinate")
        self._wb_progress2.pack(side="right")
        ttk.Label(proc, foreground="#888", wraplength=520, justify="left",
                  text="每次点击都按当前参数全量重跑，覆盖上次结果；完成后在下表更新"
                       "每位收件人的文件路径。只生成文件，Outlook 草稿在「预览与生成」步骤。"
                  ).pack(anchor="w", padx=8, pady=(0, 4))

        mframe = ttk.Frame(proc)
        mframe.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        mcols = ["name", "files", "paths", "state"]
        mheads = ["收件人名称", "将收到的文件", "文件路径", "状态"]
        self._wb_map_tree = ttk.Treeview(mframe, columns=mcols, show="headings",
                                          height=10, selectmode="browse")
        for c, h in zip(mcols, mheads):
            self._wb_map_tree.heading(c, text=h)
            self._wb_map_tree.column(
                c, width=130 if c == "name" else (180 if c in ("files", "paths") else 64))
        mvsb = ttk.Scrollbar(mframe, orient="vertical", command=self._wb_map_tree.yview)
        mhsb = ttk.Scrollbar(mframe, orient="horizontal", command=self._wb_map_tree.xview)
        self._wb_map_tree.configure(yscrollcommand=mvsb.set, xscrollcommand=mhsb.set)
        self._wb_map_tree.grid(row=0, column=0, sticky="nsew")
        mvsb.grid(row=0, column=1, sticky="ns")
        mhsb.grid(row=1, column=0, sticky="ew")
        mframe.rowconfigure(0, weight=1)
        mframe.columnconfigure(0, weight=1)
        self._wb_map_tree.bind("<Double-1>", self._wb_map_edit)
        save_row = ttk.Frame(proc)
        save_row.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Button(save_row, text="保存自定义编辑",
                   command=self._wb_map_save_edits).pack(side="left")
        ttk.Label(save_row, foreground="#888",
                  text="双击「文件路径」可手动改；保存后写入事件 process_overrides。"
                  ).pack(side="left", padx=8)

        self._wb_render_mapping_table()
        self._wb_fill_pipe_table()
        if not self._wb_pipeline_items():
            ttk.Label(self._wb_params_frame,
                      text="还没有条目。点上面的「＋ Word-MailMerge」或「＋ 附加文件」开始。",
                      style="Sub.TLabel").pack(anchor="w", padx=4, pady=6)
            return
        iid = self._wb_selected_item_id()
        item = self._wb_item_by_id(iid) if iid else None
        if item is None:
            item = self._wb_pipeline_items()[0]
            self._wb_sel_item = item.get("id")
        self._wb_render_params(item)

    def _wb_render_params(self, item):
        f = getattr(self, "_wb_params_frame", None)
        if f is None:
            return
        for c in f.winfo_children():
            c.destroy()
        # 重建前清掉上一批 tk 变量的保活列表
        self._wb_keep = []
        if item is None:
            return
        self._wb_item_params(f, item)

    def _wb_keep_var(self, var):
        """保住 tk 变量对象。

        坑：tk.StringVar 的 Python 侧一旦被 GC，它的 __del__ 会把 Tcl 变量一起
        删掉，Entry 就会莫名其妙变空（控件本身还活着，但值没了）。
        所以凡是给控件用的变量，都要在这份列表里吊住。
        """
        keep = getattr(self, "_wb_keep", None)
        if keep is None:
            keep = []
            self._wb_keep = keep
        keep.append(var)
        return var

    # ---------- 选中 / 拖动排序 / 多选删除 ----------

    def _wb_on_pipe_select(self, _evt=None):
        tree = getattr(self, "_wb_pipe_tree", None)
        if tree is None:
            return
        sel = tree.selection()
        if not sel:
            return
        self._wb_sel_item = sel[0]
        item = self._wb_item_by_id(sel[0])
        if item is not None:
            self._wb_render_params(item)

    def _wb_on_pipe_double(self, event):
        tree = getattr(self, "_wb_pipe_tree", None)
        if tree is None:
            return
        if tree.identify_column(event.x) == "#7":
            return
        iid = tree.identify_row(event.y)
        item = self._wb_item_by_id(iid) if iid else None
        if item is None:
            return
        # v3.4.6：双击不再弹重命名窗口，直接选中并聚焦左侧参数区名称输入框
        if iid:
            tree.selection_set(iid)
            tree.focus(iid)
        self._wb_sel_item = item.get("id")
        self._wb_render_params(item)
        self.root.after_idle(self._wb_focus_item_name)

    def _wb_pipe_press(self, event):
        tree = getattr(self, "_wb_pipe_tree", None)
        if tree is None:
            return
        self._wb_drag_iid = ""
        self._wb_drag_start_y = event.y
        self._wb_drag_moved = False
        self._wb_drop_iid = ""
        self._wb_drop_idx = None
        # 点「启用」列 = 开关启用状态（不进入拖动）
        if tree.identify_column(event.x) == "#7":
            iid = tree.identify_row(event.y)
            item = self._wb_item_by_id(iid) if iid else None
            if item is not None:
                item["enabled"] = not item.get("enabled", True)
                self._wb_save()
                self.root.after_idle(self._wb_after_toggle, item.get("id"))
            return
        # Ctrl/Shift = 多选，不启动拖动
        if (event.state & 0x0004) or (event.state & 0x0001):
            return
        iid = tree.identify_row(event.y)
        if iid:
            self._wb_drag_iid = iid

    def _wb_after_toggle(self, iid):
        self._wb_fill_pipe_table(keep_sel=iid)
        item = self._wb_item_by_id(iid)
        if item is not None:
            self._wb_render_params(item)

    def _wb_pipe_motion(self, event):
        tree = getattr(self, "_wb_pipe_tree", None)
        if tree is None:
            return
        iid = getattr(self, "_wb_drag_iid", "")
        if not iid:
            return
        if not self._wb_drag_moved:
            if abs(event.y - self._wb_drag_start_y) < 6:
                return
            self._wb_drag_moved = True
        rows = [r for r in tree.get_children() if not r.startswith("__drop__")]
        if iid not in rows:
            return
        target = tree.identify_row(event.y)
        idx = len(rows)
        if target and target in rows:
            idx = rows.index(target)
            bbox = tree.bbox(target)
            if bbox and event.y > bbox[1] + bbox[3] / 2.0:
                idx += 1
        self._wb_drop_idx = idx
        old = getattr(self, "_wb_drop_iid", "")
        if old:
            try:
                tree.delete(old)
            except Exception:
                pass
        self._wb_drop_iid = tree.insert("", idx,
                                         values=("↓ 放到这里 ↓", "", "", "", "", "", ""))
        tree.see(self._wb_drop_iid)

    def _wb_pipe_release(self, _event):
        tree = getattr(self, "_wb_pipe_tree", None)
        iid = getattr(self, "_wb_drag_iid", "")
        moved = getattr(self, "_wb_drag_moved", False)
        drop_iid = getattr(self, "_wb_drop_iid", "")
        idx = getattr(self, "_wb_drop_idx", None)
        self._wb_drag_iid = ""
        self._wb_drag_moved = False
        self._wb_drop_iid = ""
        self._wb_drop_idx = None
        if tree is None or not iid or not moved or idx is None:
            return
        if drop_iid:
            try:
                tree.delete(drop_iid)
            except Exception:
                pass
        rows = [r for r in tree.get_children() if r != drop_iid]
        if iid not in rows:
            return
        di = rows.index(iid)
        pos = idx - (1 if di < idx else 0)
        pos = max(0, min(pos, len(rows) - 1))
        new_order = [r for r in rows if r != iid]
        new_order.insert(pos, iid)
        items = self._wb_pipeline_items()
        by_id = dict((it.get("id"), it) for it in items)
        ordered = [by_id[r] for r in new_order if r in by_id]
        for it in items:
            if it.get("id") not in new_order:
                ordered.append(it)
        items[:] = ordered
        self._wb_save()
        self._wb_fill_pipe_table(keep_sel=iid)
        item = self._wb_item_by_id(iid)
        if item is not None:
            self._wb_render_params(item)

    # ---------- 条目增删改 ----------

    def _wb_add_item(self, kind):
        it = new_item(kind)
        self._wb_pipeline_items().append(it)
        self._wb_sel_item = it.get("id")
        self._wb_save()
        self._wb_fill_pipe_table(keep_sel=it.get("id"))
        self._wb_render_params(it)

    def _wb_del_selected_items(self):
        tree = getattr(self, "_wb_pipe_tree", None)
        sel = list(tree.selection()) if tree is not None else []
        if not sel:
            messagebox.showinfo("删除选中",
                                "请先在表格里选中要删除的条目（Ctrl/Shift 可多选）。")
            return
        items = self._wb_pipeline_items()
        names = [it.get("name") or ITEM_KIND_LABELS.get(it.get("kind"), "条目")
                 for it in items if it.get("id") in sel]
        if not messagebox.askyesno("删除选中",
                                   "确认删除 %d 个条目？\n\n%s"
                                   % (len(sel), "\n".join(names[:12]))):
            return
        items[:] = [it for it in items if it.get("id") not in sel]
        self._wb_sel_item = None
        self._wb_save()
        self._wb_fill_pipe_table()
        self._wb_render_params(None)

    def _wb_clear_items(self):
        if not self._wb_pipeline_items():
            return
        if not messagebox.askyesno("清空", "确认删除全部流水线条目？"):
            return
        self._wb_event["pipeline"] = []
        self._wb_sel_item = None
        self._wb_save()
        self._wb_fill_pipe_table()
        self._wb_render_params(None)


    def _wb_apply_item_name(self, item, name_var):
        """v3.4.6：参数区改名字段直接生效（不再弹对话框）。"""
        name = name_var.get().strip()
        if not name:
            messagebox.showwarning("改名", "条目名称不能为空")
            return
        if name == (item.get("name") or ""):
            return
        item["name"] = name
        self._wb_save()
        self._wb_fill_pipe_table(keep_sel=item.get("id"))
        self._wb_render_params(item)

    def _wb_focus_item_name(self):
        """把焦点放到当前条目的名称输入框（双击条目行时用）。"""
        e = getattr(self, "_wb_item_name_entry", None)
        try:
            if e is not None and e.winfo_exists():
                e.focus_set()
                e.select_range(0, "end")
        except Exception:
            pass

    # ---------- 条目参数区 ----------

    def _wb_download_word_example(self):
        """生成并保存示例 Word 模板（含 6 个占位符），供 MailMerge 实测。"""
        path = filedialog.asksaveasfilename(
            title="保存示例 Word 模板", defaultextension=".docx",
            initialfile="示例通知模板.docx",
            filetypes=[("Word 文档", "*.docx")])
        if not path:
            return
        try:
            wbatch = lp_v3.default_word_batch()
            tpl = Path(path)
            wbatch.create_docx(str(tpl),
                               "Capital Call Notice\n\n"
                               "Dear {{Client_Name}},\n\n"
                               "Please note the following capital call:\n\n"
                               "  Call Amount: USD {{Call_Amount}}\n"
                               "  Unfunded Commitment: USD {{Unfunded}}\n"
                               "  Due Date: {{Due_Date}}\n\n"
                               "Please remit to {{Bank_Name}} account {{Account_No}}.\n\n"
                               "Sincerely,\nFund Manager")
            messagebox.showinfo(
                "示例模板已生成",
                "已保存：\n%s\n\n占位符：{{Client_Name}} {{Call_Amount}} "
                "{{Unfunded}} {{Due_Date}} {{Bank_Name}} {{Account_No}}\n\n"
                "测试方法：流水线加一条 Word-MailMerge → 选这个模板 → "
                "点「检测占位符」→ 点「导出要素模板」填好值 → 处理。" % tpl)
        except Exception as e:
            messagebox.showerror("生成失败", str(e))

    def _wb_field_value_for(self, item, rec, ph):
        """按条目的字段映射（缺省时用自动猜测）算出某个占位符在某人身上的值。"""
        ev = self._wb_event
        mapping = item.get("mapping") or {}
        eff = dict(mapping)
        if ph not in eff:
            src = lp_v3.auto_map_placeholder(ph, ev, rec)
            if src:
                eff[ph] = src
        fmap = lp_v3.build_field_map(rec, ev)
        lp_v3.apply_mapping_overrides(fmap, eff, rec, ev)
        return str(fmap.get(ph, "") or "")

    def _wb_export_field_template(self, item):
        """导出要素模板：列 = 模板里的占位符，行 = 本事件收件人（值按当前映射预填）。"""
        params = item.get("params") or {}
        tpl = (params.get("template_path") or "").strip()
        if not tpl or not Path(tpl).exists():
            messagebox.showwarning("导出要素模板", "请先为这个条目选择有效的 .docx 模板")
            return
        try:
            phs = wm.detect_placeholders(tpl)
        except Exception as e:
            messagebox.showerror("导出要素模板", "读取模板占位符失败：\n%s" % e)
            return
        if not phs:
            messagebox.showinfo("导出要素模板",
                                "模板里没有检测到占位符。\n支持两种写法：{{字段}} 或 ?字段?")
            return
        ev = self._wb_event
        recips = ev.get("recipients") or []
        rows = [["收件人名称"] + list(phs)]
        for rec in recips:
            rows.append([rec.get("name", "")] +
                        [self._wb_field_value_for(item, rec, ph) for ph in phs])
        widths = [30] + [18] * len(phs)
        self._write_xlsx_sheets(
            "要素模板_%s.xlsx" % sanitize_filename(ev.get("name", "事件"), 30),
            [("要素模板", rows, widths)],
            "要素模板已导出",
            "已按模板里的 %d 个占位符生成要素模板（%d 位收件人）。\n"
            "每列填好后可作为外部数据表使用。" % (len(phs), len(recips)))

    def _wb_import_field_template(self, item):
        """导入要素模板：Excel 回填「收件人名称 + 占位符列」到收件人字段。

        v3.4.6：导出要素模板之后必须能导回来。
        规则：按「收件人名称」匹配本事件收件人；每个占位符列的值写进该收件人的
        `fields[占位符名]`（空单元格不覆盖旧值）；模板里检测到的占位符自动映射到
        `fields.<占位符名>`，这样「▶ 处理」/生成草稿时就能直接取到值。
        """
        params = item.get("params") or {}
        path = filedialog.askopenfilename(
            title="导入要素模板", filetypes=[("Excel", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            rows = read_spreadsheet(path)
        except Exception as e:
            messagebox.showerror("导入要素模板", "读取失败：\n%s" % e)
            return
        if not rows or len(rows) < 2:
            messagebox.showinfo("导入要素模板", "文件里没有数据行。")
            return
        header = ["" if h is None else str(h).strip() for h in rows[0]]
        name_aliases = {a.lower() for a in HEADER_ALIASES.get("name", ())}
        name_idx = None
        for i, h in enumerate(header):
            if h.lower() in name_aliases:
                name_idx = i
                break
        if name_idx is None:
            name_idx = 0          # 导出的模板第一列固定是「收件人名称」
        col_keys = {}
        for i, h in enumerate(header):
            if i == name_idx or not h:
                continue
            key = lp_v3.placeholder_key(h)
            if key:
                col_keys[i] = key
        if not col_keys:
            messagebox.showinfo("导入要素模板",
                                "除了「收件人名称」外没有可导入的占位符列。")
            return
        ev = self._wb_event
        by_name = {str(r.get("name", "")).strip(): r
                   for r in (ev.get("recipients") or [])}
        updated, unmatched, cells = 0, [], 0
        filled_keys = set()
        for row in rows[1:]:
            name = str(row[name_idx] or "").strip() if name_idx < len(row) else ""
            if not name or name.startswith("示例："):
                continue
            rec = by_name.get(name)
            if rec is None:
                unmatched.append(name)
                continue
            fields = rec.setdefault("fields", {})
            for i, key in col_keys.items():
                val = row[i] if i < len(row) else None
                val = "" if val is None else str(val).strip()
                if val == "":
                    continue      # 空单元格不覆盖已有值
                fields[key] = val
                filled_keys.add(key)
                cells += 1
            updated += 1
        # 模板里的占位符 → 刚导入的字段来源
        mapped = 0
        tpl = (params.get("template_path") or "").strip()
        phs = []
        if tpl and Path(tpl).exists():
            try:
                phs = wm.detect_placeholders(tpl)
            except Exception:
                phs = []
        mapping = item.setdefault("mapping", {})
        imported_keys = set(col_keys.values())
        for ph in phs:
            key = lp_v3.placeholder_key(ph)
            # 只有真的有值的列才自动接映射；空列不覆盖用户已有/内置来源
            if key in imported_keys and key in filled_keys:
                mapping[ph] = "fields." + key
                mapped += 1
        self._wb_save()
        self._wb_render_params(item)
        msg = "已导入 %d 位收件人的要素（%d 个单元格）。" % (updated, cells)
        if mapped:
            msg += "\n已自动把 %d 个占位符接到对应字段来源。" % mapped
        if unmatched:
            msg += "\n\n以下收件人名称在事件里找不到（未导入）：\n" + "、".join(unmatched[:10])
            if len(unmatched) > 10:
                msg += "…"
        messagebox.showinfo("要素模板已导入", msg)

    def _wb_export_attach_path_template(self, item):
        """导出路径模板：收件人 + 该收件人在这条附加文件条目下会取到的路径。"""
        ev = self._wb_event
        rows = [["收件人名称", "附件路径"]]
        for rec in ev.get("recipients") or []:
            unit = {"rep": rec, "recipients": [rec],
                    "lp_name": rec.get("name", ""), "group_key": rec.get("name", ""),
                    "group": (rec.get("group") or "").strip()}
            try:
                paths, _label = lp_v3.resolve_item_paths(item, unit, ev)
            except Exception:
                paths = []
            rows.append([rec.get("name", ""), ";".join(str(p) for p in paths)])
        self._write_xlsx_sheets(
            "附件路径模板_%s.xlsx" % sanitize_filename(ev.get("name", "事件"), 30),
            [("附件路径", rows, [30, 80])],
            "路径模板已导出",
            "已导出 %d 位收件人的附件路径。\n"
            "手填路径条目所有人相同；按「收件人附件/字段/外部表」取值的会逐人不同。"
            % (len(rows) - 1))

    def _wb_item_params(self, parent, item):
        kind = item.get("kind") or "generate"
        params = item.setdefault("params", {})
        title = item.get("name") or ITEM_KIND_LABELS.get(kind, "条目")
        box = ttk.LabelFrame(parent, text=" 选中条目：%s " % title,
                             style="Card.TLabelframe")
        box.pack(fill="x", pady=(0, 6))

        r0 = ttk.Frame(box)
        r0.pack(fill="x", padx=8, pady=(6, 2))
        ttk.Label(r0, text="名称：").pack(side="left")
        name_var = self._wb_keep_var(tk.StringVar(value=item.get("name") or ""))
        name_entry = ttk.Entry(r0, textvariable=name_var, width=28)
        name_entry.pack(side="left", padx=4)
        self._wb_item_name_entry = name_entry
        name_entry.bind("<Return>", lambda _e: self._wb_apply_item_name(item, name_var))
        ttk.Button(r0, text="改名", width=6,
                   command=lambda: self._wb_apply_item_name(item, name_var)).pack(side="left")
        ttk.Label(r0, text="   类型：%s" % ITEM_KIND_LABELS.get(kind, ""),
                  foreground="#666").pack(side="left", padx=(12, 0))
        ToolTip(r0, ITEM_KIND_DESC.get(kind, ""))

        if kind == "generate":
            r1 = ttk.Frame(box)
            r1.pack(fill="x", padx=8, pady=4)
            ttk.Label(r1, text="Word 模板：").pack(side="left")
            tpl_var = self._wb_keep_var(tk.StringVar(value=params.get("template_path", "")))
            ttk.Entry(r1, textvariable=tpl_var, width=32).pack(side="left", padx=4)
            ttk.Button(r1, text="选择…", width=6,
                       command=lambda: self._wb_pick_template(item, tpl_var)).pack(side="left")
            ttk.Button(r1, text="检测占位符", width=10,
                       command=lambda: self._wb_detect_placeholders(item)).pack(side="left", padx=4)
            # v3.4.6：不再弹「说明」窗口，关键信息直接内联在参数区
            ttk.Label(box, justify="left", foreground="#888", wraplength=490,
                      text="占位符写法：{{字段}} 或 ?字段?（写在 Word 模板里）。\n"
                           "处理顺序固定：转 PDF → 加水印 → 加密；"
                           "水印文字与各人打开口令在「收件人」处维护。"
                      ).pack(anchor="w", padx=8, pady=(2, 0))
            hrow = ttk.Frame(box)
            hrow.pack(fill="x", padx=8, pady=(2, 4))
            ttk.Button(hrow, text="导出要素模板", width=12,
                       command=lambda: self._wb_export_field_template(item)
                       ).pack(side="left")
            ttk.Button(hrow, text="导入要素模板", width=12,
                       command=lambda: self._wb_import_field_template(item)
                       ).pack(side="left", padx=(4, 0))
            ttk.Button(hrow, text="下载示例 Word 模板", width=17,
                       command=self._wb_download_word_example).pack(side="left", padx=(4, 0))
        else:

            r1 = ttk.Frame(box)
            r1.pack(fill="x", padx=8, pady=4)
            ttk.Label(r1, text="文件（每行一个路径）：").pack(anchor="w")
            # v3.4.5：文本框与按钮改为上下两行，默认窗口宽度下按钮不会被裁掉
            files_text = tk.Text(r1, height=4, width=40, wrap="none")
            files_text.pack(fill="x", padx=0, pady=(2, 4))
            for p in (params.get("files") or []):
                files_text.insert("end", str(p) + "\n")
            btns = ttk.Frame(r1)
            btns.pack(fill="x")
            b_add = ttk.Button(btns, text="添加文件…", width=12,
                               command=lambda: self._wb_add_attach_files(item, files_text))
            b_add.pack(side="left")
            b_imp = ttk.Button(btns, text="从 Excel 导入路径…", width=18,
                               command=lambda: self._wb_import_attach_paths(item, files_text))
            b_imp.pack(side="left", padx=4)
            b_save = ttk.Button(btns, text="保存", width=6,
                                command=lambda: self._wb_save_attach_files(item, files_text))
            b_save.pack(side="left")
            b_tpl = ttk.Button(btns, text="导出路径模板", width=12,
                               command=lambda: self._wb_export_attach_path_template(item))
            b_tpl.pack(side="left", padx=4)
            ttk.Label(box, text=(
                "每行一个文件路径。从 Excel 导入时路径与收件人一一对应；"
                "所有人同一份文件用「统一附件」按钮设置。"
                if False else
                "每行一个文件路径，从 Excel 导入时与收件人一一对应；"
                "全员共用同一份文件请用「统一附件」按钮。"),
                foreground="#888", wraplength=880, justify="left").pack(anchor="w", padx=8, pady=(0, 4))

        # --- 处理勾选（顺序固定：转 PDF → 加水印 → 加密）---
        r2 = ttk.Frame(box)
        r2.pack(fill="x", padx=8, pady=4)
        ttk.Label(r2, text="处理：").pack(side="left")
        to_pdf_var = self._wb_keep_var(tk.BooleanVar(value=bool(params.get("to_pdf"))))
        wm_var = self._wb_keep_var(tk.BooleanVar(value=bool(params.get("watermark"))))
        enc_var = self._wb_keep_var(tk.BooleanVar(value=bool(params.get("encrypt"))))
        if kind == "generate":
            ttk.Checkbutton(r2, text="转成 PDF", variable=to_pdf_var,
                            command=lambda: self._wb_item_flag(item, "to_pdf", to_pdf_var)
                            ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(r2, text="加水印", variable=wm_var,
                        command=lambda: self._wb_item_flag(item, "watermark", wm_var)
                        ).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(r2, text="加密（AES-256）", variable=enc_var,
                        command=lambda: self._wb_item_flag(item, "encrypt", enc_var)
                        ).pack(side="left")
        ToolTip(r2, "处理顺序固定，你不用排：转 PDF → 加水印 → 加密。\n"
                   "加密只对 PDF 生效；打开口令取该收件人的打开口令。")
        ttk.Label(box, foreground="#888",
                  text="水印文字与各人打开口令在「收件人」处维护",
                  wraplength=500, justify="left").pack(anchor="w", padx=8, pady=(0, 2))

        # --- 加密选项：发送人统一密码（即 GP 统一权限口令机制，方便发送人核对）---
        if params.get("encrypt"):
            r_gp = ttk.Frame(box)
            r_gp.pack(fill="x", padx=8, pady=(0, 4))
            _ev_opts = (self._wb_event or {}).get("options") or {}
            gp_mode_var = self._wb_keep_var(tk.BooleanVar(
                value=bool(_ev_opts.get("gp_owner_mode",
                                        self.settings.get("gp_owner_mode")))))
            gp_pw_var = self._wb_keep_var(tk.StringVar(
                value=_ev_opts.get("gp_owner_pw", self.settings.get("gp_owner_pw", ""))))

            def _gp_changed(mode_v=gp_mode_var, pw_v=gp_pw_var):
                mode = bool(mode_v.get())
                pw = pw_v.get().strip()
                self.settings["gp_owner_mode"] = mode
                self.settings["gp_owner_pw"] = pw
                self._persist_settings()        # v3.4.4：原为不存在的 save_data()
                # 引擎/预检只读 event["options"]，这里必须同步过去，否则勾了等于没勾
                ev = self._wb_event
                if ev is not None:
                    eopts = ev.setdefault("options", {})
                    eopts["gp_owner_mode"] = mode
                    eopts["gp_owner_pw"] = pw
                    self._wb_save()

            # v3.4.4：输入口令时 StringVar 一变化就保存（真实键入 / 测试都可靠）
            gp_pw_var.trace_add("write", lambda *_a: _gp_changed())

            ttk.Checkbutton(r_gp, text="发送人统一密码（所有文件同一口令，便于发送人核对）",
                            variable=gp_mode_var,
                            command=_gp_changed).pack(side="left")
            e_gp = ttk.Entry(r_gp, textvariable=gp_pw_var, width=20, show="*")
            e_gp.pack(side="left", padx=6)
            for _seq in ("<KeyRelease>", "<Return>", "<FocusOut>"):
                e_gp.bind(_seq, lambda _e: _gp_changed())
            ToolTip(e_gp, "统一打开口令（收件人口令之外，发送人自己核对用）。"
                          "不能与任何收件人的打开口令相同。留空则生成时自动创建。")

        # --- 权限位（仅加密时出现）---
        if params.get("encrypt"):
            r3 = ttk.Frame(box)
            r3.pack(fill="x", padx=8, pady=(0, 4))
            ttk.Label(r3, text="权限：").pack(side="left")
            perms = params.get("permissions")
            perms = list(perms) if isinstance(perms, list) else list(lp_v3.DEFAULT_PERMISSIONS)
            for pk in lp_v3.PERMISSIONS:
                pv = self._wb_keep_var(tk.BooleanVar(value=pk in perms))
                ttk.Checkbutton(r3, text=lp_v3.PERMISSION_LABELS.get(pk, pk),
                                variable=pv,
                                command=lambda k=pk, v=pv: self._wb_item_perm(item, k, v)
                                ).pack(side="left", padx=(0, 10))
            ttk.Label(box, foreground="#888", text=(
                "默认只允许打印。全部取消 = 收件人只能打开查看，不能打印/编辑/复制。"),
                wraplength=500, justify="left").pack(anchor="w", padx=8, pady=(0, 2))

        # --- 产出粒度（条目级：这一条产出几份文件）---
        r4 = ttk.Frame(box)
        r4.pack(fill="x", padx=8, pady=4)
        ttk.Label(r4, text="产出颗粒度（这一条产出几份文件）：").pack(side="left")
        cur_gran = item.get("granularity") or "per_lp"
        gran_var = self._wb_keep_var(tk.StringVar(value=GRANULARITY_LABELS.get(cur_gran, cur_gran)))
        cb = ttk.Combobox(r4, textvariable=gran_var, state="readonly", width=16,
                          values=[GRANULARITY_LABELS[g] for g in lp_v3.GRANULARITIES])
        cb.pack(side="left", padx=4)
        cb.bind("<<ComboboxSelected>>",
                lambda _e: self._wb_set_item_gran(item, gran_var.get()))
        ToolTip(cb, "\n".join("· %s：%s" % (GRANULARITY_LABELS[g],
                                            lp_v3.GRANULARITY_DESC.get(g, ""))
                              for g in lp_v3.GRANULARITIES)
                   + "\n\n注意：事件里的「发送模式」决定几封邮件，这里决定这一条产出几份文件。")
        ttk.Label(r4, foreground="#888",
                  text=lp_v3.GRANULARITY_DESC.get(cur_gran, "")).pack(side="left", padx=8)

        if kind == "generate":
            mbox = ttk.LabelFrame(parent, text=" 字段映射（Word 占位符 → 数据来源） ",
                                  style="Card.TLabelframe")
            mbox.pack(fill="both", expand=True, pady=(0, 6))
            self._wb_mapping_panel(mbox, item)

    # ---------- 字段级来源（映射表）----------

    def _wb_sheet_cols(self, item):
        path = ((item.get("sheet") or {}).get("path") or "").strip()
        if not path or not Path(path).exists():
            return []
        try:
            rows = lp_v3.sheet_to_dicts(lp_v3.read_spreadsheet(path))
        except Exception:
            return []
        return list(rows[0].keys()) if rows else []

    def _wb_sheet_row(self, item, rec):
        if not rec:
            return None
        try:
            ctx = lp_v3.RunContext(self._wb_event, Path.cwd())
            return ctx.sheet_for(item, rec)
        except Exception:
            return None

    def _wb_field_source_options(self, item, rec0):
        return lp_v3.field_source_options(self._wb_event, rec0,
                                         self._wb_sheet_cols(item))

    def _wb_src_label(self, src, opts):
        if not src:
            return "（未映射）"
        for key, cat, label in opts:
            if key == src:
                return "%s：%s" % (cat, label)
        return src

    def _wb_mapping_panel(self, parent, item):
        ev = self._wb_event
        recips = ev.get("recipients") or []
        rec0 = recips[0] if recips else {}
        mapping = item.setdefault("mapping", {})
        opts = self._wb_field_source_options(item, rec0)
        sheet_row = self._wb_sheet_row(item, rec0)
        rows = lp_v3.item_placeholders(item, ev, rec0, sheet_row)

        if not mapping:
            ttk.Label(parent, text=(
                "还没有映射。点上面参数区的「检测占位符」自动读取模板里的占位符，"
                "再逐个确认来源（双击表格行可改）。"),
                style="Sub.TLabel").pack(anchor="w", padx=8, pady=6)

        cols = ["ph", "src", "val"]
        heads = ["Word 占位符", "填什么（数据来源）",
                 "预览值（以「%s」为例）" % (rec0.get("name") or "第一位收件人")]
        tframe = ttk.Frame(parent)
        tframe.pack(fill="both", expand=True, padx=8, pady=4)
        tree = ttk.Treeview(tframe, columns=cols, show="headings", height=7,
                            selectmode="extended")
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
        tree.column("ph", width=170)
        tree.column("src", width=190)
        tree.column("val", width=180)
        tvsb = ttk.Scrollbar(tframe, orient="vertical", command=tree.yview)
        thsb = ttk.Scrollbar(tframe, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=tvsb.set, xscrollcommand=thsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        tvsb.grid(row=0, column=1, sticky="ns")
        thsb.grid(row=1, column=0, sticky="ew")
        tframe.rowconfigure(0, weight=1)
        tframe.columnconfigure(0, weight=1)
        self._wb_map_iids = {}
        for i, (ph, src, val) in enumerate(rows):
            iid = "m%d" % i
            self._wb_map_iids[iid] = ph
            tree.insert("", "end", iid=iid,
                        values=(ph, self._wb_src_label(src, opts), val))
        # v3.4.8：双击哪一行就编辑哪个占位符（原来永远传 item、弹窗固定第一个）
        tree.bind("<Double-1>", lambda e, it=item: self._wb_edit_mapping(
            it, (getattr(self, "_wb_map_iids", None) or {}).get(tree.identify_row(e.y))))

        bar = ttk.Frame(parent)
        bar.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Button(bar, text="＋ 添加映射行",
                   command=lambda: self._wb_add_mapping(item)).pack(side="left")
        ttk.Button(bar, text="删除选中",
                   command=lambda: self._wb_del_mapping(item, tree)).pack(side="left", padx=4)
        ttk.Label(bar, foreground="#888", text=(
            "权限口令不是可映射的字段（架构上无法写进文档或邮件）。")
        ).pack(side="left", padx=8)

        # 外部数据表：只有在真的选了「外部数据表」来源时才出现
        if any(str(v).startswith("sheet.") for v in mapping.values()):
            self._wb_sheet_panel(parent, item)

    def _wb_sheet_panel(self, parent, item):
        """外部数据表配置面板。

        生成文档条目选了 sheet. 来源、附加文件条目选了外部数据表列来源时共用。
        """
        sh = item.setdefault("sheet", {"path": "", "key_col": "", "match_col": "name"})
        sbox = ttk.LabelFrame(parent, text=" 外部数据表 ", style="Card.TLabelframe")
        sbox.pack(fill="x", padx=8, pady=(0, 4))
        s1 = ttk.Frame(sbox)
        s1.pack(fill="x", padx=8, pady=4)
        ttk.Label(s1, text="数据表：").pack(side="left")
        path_var = self._wb_keep_var(tk.StringVar(value=sh.get("path", "")))
        ttk.Entry(s1, textvariable=path_var, width=44).pack(side="left", padx=4)
        ttk.Button(s1, text="选择…", width=6,
                   command=lambda: self._wb_pick_sheet(item)).pack(side="left")
        ttk.Label(s1, text="   匹配列（表内）：").pack(side="left", padx=(10, 0))
        key_var = self._wb_keep_var(tk.StringVar(value=sh.get("key_col", "")))
        ttk.Entry(s1, textvariable=key_var, width=16).pack(side="left", padx=4)
        ttk.Label(s1, text="   收件人匹配字段：").pack(side="left", padx=(10, 0))
        match_var = self._wb_keep_var(tk.StringVar(value=sh.get("match_col", "name")))
        ttk.Entry(s1, textvariable=match_var, width=16).pack(side="left", padx=4)
        ttk.Label(s1, foreground="#888", text="（默认 name = 收件人名称）").pack(side="left", padx=6)
        s2 = ttk.Frame(sbox)
        s2.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Button(s2, text="保存数据表配置", width=14, command=lambda: self._wb_save_sheet_cfg(
            item, path_var.get(), key_var.get(), match_var.get())).pack(side="left")
        ttk.Button(s2, text="下载数据表模板", width=14,
                   command=self._wb_export_sheet_template).pack(side="left", padx=4)
        ttk.Button(s2, text="校验数据表", width=12,
                   command=lambda: self._wb_check_sheet(item)).pack(side="left")
        ttk.Label(s2, foreground="#666", text=(
            "第一列默认当匹配列；导入时会校验「列是否存在 / key 是否重复 / "
            "有多少收件人匹配不上」。")
        ).pack(side="left", padx=8)

    def _wb_add_mapping(self, item):
        mapping = item.setdefault("mapping", {})
        i = len(mapping) + 1
        name = "{{字段%d}}" % i
        while name in mapping:
            i += 1
            name = "{{字段%d}}" % i
        mapping[name] = ""
        self._wb_save()
        self._wb_render_params(item)

    def _wb_del_mapping(self, item, tree):
        sel = list(tree.selection())
        if not sel:
            return
        mapping = item.setdefault("mapping", {})
        iids = getattr(self, "_wb_map_iids", {}) or {}
        for iid in sel:
            ph = iids.get(iid)
            if ph:
                mapping.pop(ph, None)
        self._wb_save()
        self._wb_render_params(item)

    def _wb_edit_mapping(self, item, initial_ph=None):
        """占位符来源弹窗。

        v3.4.8 修复：
          · 占位符下拉必须带 values（原来没传 → 只能看到第一个/无法选其他）；
          · 切换占位符时，来源下拉要跟着刷新成该占位符当前的映射；
          · 双击映射表哪一行，就预选那一行的占位符（原来永远第一个）。
        """
        ev = self._wb_event
        recips = ev.get("recipients") or []
        rec0 = recips[0] if recips else {}
        mapping = item.setdefault("mapping", {})
        if not mapping:
            return
        opts = self._wb_field_source_options(item, rec0)
        src_keys = [k for k, _c, _l in opts]
        src_labels = ["%s：%s" % (c, l) for _k, c, l in opts]
        ph_list = list(mapping.keys())
        initial = initial_ph if initial_ph in mapping else ph_list[0]

        win = tk.Toplevel(self.root)
        win.title("占位符来源")
        win.grab_set()
        ttk.Label(win, text="选择占位符，再选它的数据来源：", justify="left").pack(
            anchor="w", padx=12, pady=(10, 4))
        ph_var = tk.StringVar(value=initial)
        ph_box = ttk.Combobox(win, textvariable=ph_var, state="readonly", width=40,
                              values=ph_list)
        ph_box.pack(anchor="w", padx=12)
        ttk.Label(win, text="来源：", justify="left").pack(anchor="w", padx=12, pady=(10, 4))
        src_var = tk.StringVar(value="")
        src_box = ttk.Combobox(win, textvariable=src_var, state="readonly", width=40,
                               values=src_labels)
        src_box.pack(anchor="w", padx=12)

        def _src_label_of(ph):
            cur = mapping.get(ph, "")
            idx = src_keys.index(cur) if cur in src_keys else 0
            return src_labels[idx] if src_labels else ""

        def _sync_source(_evt=None):
            src_var.set(_src_label_of(ph_var.get()))

        ph_box.bind("<<ComboboxSelected>>", _sync_source)
        _sync_source()

        ttk.Label(win, foreground="#666", justify="left", wraplength=430, text=(
            "· 收件人名称 / 水印名称 / 打开口令 / 日期 —— 内置\n"
            "· 收件人字段 —— 收件人表里你自己加的列（可临时加列）\n"
            "· 事件常量 —— 事件选项里定义的固定值\n"
            "· 外部数据表 —— 从 Excel 按收件人匹配进来的列")
        ).pack(anchor="w", padx=12, pady=(6, 0))

        def ok():
            sel = src_box.get()
            key = ""
            for i, s in enumerate(src_labels):
                if s == sel:
                    key = src_keys[i]
                    break
            mapping[ph_var.get()] = key
            self._wb_save()
            win.destroy()
            self._wb_render_params(item)

        btns = ttk.Frame(win)
        btns.pack(pady=10)
        ttk.Button(btns, text="确定", command=ok).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=6)

    def _wb_guess_source(self, ph, ev, rec0, sheet_cols):
        key = lp_v3.placeholder_key(ph)
        for c in sheet_cols:
            if c.lower() == key.lower():
                return "sheet." + c
        return lp_v3.auto_map_placeholder(ph, ev, rec0)

    def _wb_detect_placeholders(self, item):
        tpl = (item.get("params") or {}).get("template_path", "").strip()
        if not tpl or not Path(tpl).exists():
            messagebox.showwarning("检测占位符", "请先为这个条目选择有效的 .docx 模板")
            return
        try:
            found = wm.detect_placeholders(tpl)
        except Exception as e:
            messagebox.showerror("检测失败", str(e))
            return
        if not found:
            messagebox.showinfo("检测结果", "模板里没有检测到占位符。\n"
                                "支持两种写法：{{字段}} 或 ?字段?")
            return
        ev = self._wb_event
        recips = ev.get("recipients") or []
        rec0 = recips[0] if recips else {}
        sheet_cols = self._wb_sheet_cols(item)
        mapping = item.setdefault("mapping", {})
        added = 0
        for ph in found:
            if ph not in mapping:
                mapping[ph] = self._wb_guess_source(ph, ev, rec0, sheet_cols)
                added += 1
        self._wb_save()
        self._wb_render_params(item)
        unresolved = [p for p, s in mapping.items() if not s]
        msg = "检测到 %d 个占位符，新增 %d 个映射。\n" % (len(found), added)
        if unresolved:
            msg += ("\n以下占位符没有找到对应字段（预检会报错），请双击手动指定来源：\n"
                    + "\n".join(unresolved[:10]))
        messagebox.showinfo("检测结果", msg)

    # ---------- 模板 / 文件选择 ----------

    def _wb_pick_template(self, item, var):
        path = filedialog.askopenfilename(title="选择 Word 模板",
                                          filetypes=[("Word 文档", "*.docx;*.doc")])
        if path:
            var.set(path)
            item.setdefault("params", {})["template_path"] = path
            self._wb_save()
            self._wb_render_params(item)

    def _wb_attach_files_from_text(self, tw):
        out = []
        for line in tw.get("1.0", "end").splitlines():
            s = line.strip()
            if s and s not in out:
                out.append(s)
        return out

    def _wb_write_attach_files(self, item, files, tw):
        item.setdefault("params", {})["files"] = files
        self._wb_save()
        tw.delete("1.0", "end")
        for p in files:
            tw.insert("end", p + "\n")
        self._wb_fill_pipe_table(keep_sel=item.get("id"))

    def _wb_add_attach_files(self, item, tw):
        paths = filedialog.askopenfilenames(title="选择要附加的文件")
        if not paths:
            return
        cur = self._wb_attach_files_from_text(tw)
        for p in paths:
            if p not in cur:
                cur.append(p)
        self._wb_write_attach_files(item, cur, tw)

    def _wb_path_preview(self, item, rec):
        """路径来源预览：告诉用户「以第一位收件人为例会取到什么」。"""
        src = str((item.get("params") or {}).get("path_source") or lp_v3.PATH_SOURCE_MANUAL)
        who = (rec or {}).get("name") or "第一位收件人"
        if src == lp_v3.PATH_SOURCE_MANUAL:
            files = (item.get("params") or {}).get("files") or []
            if not files:
                return "（当前：手填路径，尚未添加文件）"
            return "（当前：手填 %d 个路径，所有收件人一样）" % len(files)
        if src == lp_v3.PATH_SOURCE_RECIPIENT:
            atts = (rec or {}).get("attachments") or []
            if not atts:
                return "（以「%s」为例：该收件人还没填附件路径）" % who
            return "（以「%s」为例：%s）" % (who, "、".join(Path(p).name for p in atts))
        if src.startswith(lp_v3.PATH_SOURCE_FIELD_PREFIX):
            k = src[len(lp_v3.PATH_SOURCE_FIELD_PREFIX):]
            v = ((rec or {}).get("fields") or {}).get(k, "")
            return "（以「%s」为例：字段「%s」= %s）" % (who, k, v or "（空）")
        if src.startswith(lp_v3.PATH_SOURCE_SHEET_PREFIX):
            col = src[len(lp_v3.PATH_SOURCE_SHEET_PREFIX):]
            return "（外部数据表「%s」列，需在下方配置数据表）" % col
        return ""

    def _wb_set_item_path_source(self, item, label, opts):
        """切换附加文件条目的路径来源。"""
        for k, _g, l in opts:
            if l == label:
                params = item.setdefault("params", {})
                params["path_source"] = k
                self._wb_save()
                self._wb_render_params(item)
                return

    def _wb_save_attach_files(self, item, tw):
        self._wb_write_attach_files(item, self._wb_attach_files_from_text(tw), tw)

    def _wb_import_attach_paths(self, item, tw):
        path = filedialog.askopenfilename(title="选择路径清单（Excel/CSV）",
                                          filetypes=[("表格", "*.xlsx;*.csv"),
                                                     ("所有文件", "*.*")])
        if not path:
            return
        try:
            rows = lp_v3.read_spreadsheet(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not rows:
            messagebox.showerror("读取失败", "文件里读不到任何行")
            return
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        win = tk.Toplevel(self.root)
        win.title("选择路径列")
        win.grab_set()
        ttk.Label(win, text="哪一列是文件路径？", justify="left").pack(
            anchor="w", padx=12, pady=(10, 4))
        var = tk.StringVar(value=headers[0] if headers else "")
        ttk.Combobox(win, textvariable=var, state="readonly", width=36,
                     values=[h for h in headers if h]).pack(anchor="w", padx=12)

        def ok():
            win.destroy()
            col = var.get()
            if not col or col not in headers:
                return
            ci = headers.index(col)
            cur = self._wb_attach_files_from_text(tw)
            added = 0
            for r in rows[1:]:
                if ci >= len(r):
                    continue
                for p in lp_v3.split_paths(str(r[ci] or "")):
                    if p and p not in cur:
                        cur.append(p)
                        added += 1
            self._wb_write_attach_files(item, cur, tw)
            messagebox.showinfo("导入完成",
                                "从「%s」的「%s」列导入 %d 条路径。"
                                % (Path(path).name, col, added))

        btns = ttk.Frame(win)
        btns.pack(pady=10)
        ttk.Button(btns, text="导入", command=ok).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=win.destroy).pack(side="left", padx=6)

    # ---------- 外部数据表 ----------

    def _wb_pick_sheet(self, item):
        path = filedialog.askopenfilename(title="选择外部数据表",
                                          filetypes=[("表格", "*.xlsx;*.csv"),
                                                     ("所有文件", "*.*")])
        if not path:
            return
        sh = item.setdefault("sheet", {"path": "", "key_col": "", "match_col": "name"})
        sh["path"] = path
        if not sh.get("key_col"):
            try:
                rows = lp_v3.read_spreadsheet(path)
                if rows and rows[0]:
                    sh["key_col"] = str(rows[0][0]).strip()
            except Exception:
                pass
        self._wb_save()
        self._wb_render_params(item)
        self._wb_check_sheet(item)

    def _wb_save_sheet_cfg(self, item, path, key_col, match_col):
        sh = item.setdefault("sheet", {"path": "", "key_col": "", "match_col": "name"})
        sh["path"] = path.strip()
        sh["key_col"] = key_col.strip()
        sh["match_col"] = match_col.strip() or "name"
        self._wb_save()
        self._wb_render_params(item)

    def _wb_export_sheet_template(self):
        headers = ["收件人名称", "Call_Amount", "Unfunded", "Due_Date"]
        row = ["示例：苏州某某科技有限公司", "200,000", "800,000", "2026-11-15"]
        guide = [
            ["列名", "必填", "说明"],
            ["第一列（示例为「收件人名称」）", "是",
             "匹配列。必须能与收件人的名称完全对上（默认取收件人名称字段）"],
            ["其它任意列", "否",
             "数据列。在「字段映射」里把 Word 占位符指向这些列（来源选「外部数据表」）"],
        ]
        self._write_xlsx_sheets(
            "外部数据表模板.xlsx",
            [("数据表", [headers, row], [36, 16, 16, 16]),
             ("填写说明", guide, [30, 8, 74])],
            "模板已导出",
            "数据表模板已导出。\n第一列必须与收件人名称能对上；\n"
            "其余列会在字段映射的「外部数据表」来源里变成可选字段。")

    def _wb_check_sheet(self, item):
        sh = item.get("sheet") or {}
        path = sh.get("path", "")
        if not path:
            return
        try:
            rows = lp_v3.read_spreadsheet(path)
        except Exception as e:
            messagebox.showerror("数据表读取失败", str(e))
            return
        if not rows:
            messagebox.showerror("数据表校验", "文件是空的（读不到任何行）")
            return
        headers = [str(h).strip() if h is not None else "" for h in rows[0]]
        key_col = (sh.get("key_col") or (headers[0] if headers else "")).strip()
        problems, warns = [], []
        if key_col not in headers:
            problems.append("匹配列「%s」不在表头里。表头是：%s"
                            % (key_col, "、".join(h for h in headers if h)))
            key_col = headers[0] if headers else ""
        data = [r for r in rows[1:] if any(str(c).strip() for c in r if c is not None)]
        if not data:
            problems.append("表头下面没有数据行")
        seen, dup = {}, []
        for i, r in enumerate(data):
            try:
                v = str(r[headers.index(key_col)]).strip() if key_col in headers else ""
            except Exception:
                v = ""
            if not v:
                warns.append("第 %d 行没有填匹配列的值" % (i + 2))
            elif v in seen:
                dup.append(v)
            else:
                seen[v] = i + 2
        if dup:
            problems.append("匹配列有重复值：%s（前一个在第 %d 行）"
                            % ("、".join(dup[:5]), seen.get(dup[0], 0)))
        match_col = (sh.get("match_col") or "name").strip()
        recips = (self._wb_event or {}).get("recipients") or []
        unmatched = []
        if recips and match_col == "name":
            names = set(seen.keys())
            for r in recips:
                if r.get("name", "") not in names:
                    unmatched.append(r.get("name", ""))
        msg = "数据表：%s\n匹配列：%s\n数据行：%d" % (Path(path).name, key_col, len(data))
        if recips and match_col == "name":
            msg += "\n能匹配上的收件人：%d / %d" % (len(recips) - len(unmatched), len(recips))
        if problems:
            messagebox.showerror("数据表校验未通过",
                                 msg + "\n\n" + "\n".join("✕ " + p for p in problems))
            return
        extra = []
        if unmatched:
            extra.append("有 %d 位收件人在数据表里找不到对应行（他们会拿不到这些字段）：\n%s"
                         % (len(unmatched), "、".join(unmatched[:8])))
        if warns:
            extra.append("\n".join(warns[:5]))
        if extra:
            messagebox.showwarning("数据表校验通过（有问题需要注意）",
                                   msg + "\n\n" + "\n\n".join(extra))
        else:
            messagebox.showinfo("数据表校验通过", msg + "\n\n没有问题。")

    # ---------- 勾选项 / 粒度 ----------

    def _wb_item_flag(self, item, key, var):
        item.setdefault("params", {})[key] = bool(var.get())
        self._wb_save()
        self._wb_fill_pipe_table(keep_sel=item.get("id"))
        self._wb_render_params(item)

    def _wb_item_perm(self, item, pk, var):
        params = item.setdefault("params", {})
        perms = params.get("permissions")
        perms = list(perms) if isinstance(perms, list) else list(lp_v3.DEFAULT_PERMISSIONS)
        perms = [x for x in perms if x in lp_v3.PERMISSIONS]
        if var.get():
            if pk not in perms:
                perms.append(pk)
        else:
            perms = [x for x in perms if x != pk]
        params["permissions"] = perms
        self._wb_save()
        self._wb_render_params(item)

    def _wb_set_item_gran(self, item, label):
        for g, lab in GRANULARITY_LABELS.items():
            if lab == label:
                item["granularity"] = g
                break
        self._wb_save()
        self._wb_fill_pipe_table(keep_sel=item.get("id"))
        self._wb_render_params(item)

    # ---------- 步骤 3：附件 ----------

    def _wb_step_attachments(self, parent):
        ev = self._wb_event
        ttk.Label(parent, text="附件（快捷挂载）", style="Head.TLabel").pack(anchor="w")
        ttk.Label(parent, text=(
            "不想走流水线时的快捷入口：直接把磁盘上的文件挂给收件人。\n"
            "（要每人一份的内容，请用「文档流水线」里的「Word-MailMerge」条目）\n"
            "静态文件按范围分配：全部收件人 / 指定收件人 / 按合并组。"),
            style="Sub.TLabel", justify="left").pack(anchor="w", pady=(0, 6))

        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 4))
        ttk.Button(bar, text="添加附件", command=self._wb_add_attachment).pack(side="left")
        ttk.Button(bar, text="删除选中", command=self._wb_del_attachment).pack(side="left", padx=4)
        ttk.Button(bar, text="检查文件存在性", command=self._wb_check_attachments2
                   ).pack(side="left")

        cols = ["name", "path", "scope", "target", "ok", "go"]
        heads = ["附件名", "路径", "范围", "目标", "状态", ""]
        # 同上：Treeview 必须建在 frame 里，否则会被排到 frame 下面
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="extended")
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
            tree.column(c, width=200 if c == "path" else 110)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self._wb_att_tree = tree

        for a in ev.get("attachments", []):
            target = ""
            if a.get("scope") == "specific":
                target = "、".join(a.get("targets") or [])
            elif a.get("scope") == "group":
                target = a.get("group", "")
            exists = "存在" if (a.get("path") and Path(a["path"]).exists()) else "文件不存在"
            tree.insert("", "end", values=[
                a.get("name", ""), a.get("path", ""),
                {"all": "全部", "specific": "指定", "group": "按组"}.get(a.get("scope"), "全部"),
                target or "—", exists, "›"])
        ttk.Label(parent, foreground="#444", text=(
            "流水线产出的文件（Word→PDF→水印→加密）自动进入对应收件人的附件包，"
            "此处仅管理额外附加的静态文件。")).pack(fill="x", pady=(6, 0))

    def _wb_add_attachment(self):
        path = filedialog.askopenfilename(title="选择附件文件")
        if not path:
            return
        name = simpledialog.askstring("添加附件", "附件显示名（可空=用文件名）：",
                                      initialvalue=Path(path).name)
        if name is None:
            return
        att = normalize_attachment({"name": name.strip() or Path(path).name,
                                    "path": path, "scope": "all"})
        self._wb_event.setdefault("attachments", []).append(att)
        self._wb_save()
        self._wb_show_step(self._wb_step)

    def _wb_del_attachment(self):
        tree = getattr(self, "_wb_att_tree", None)
        ev = self._wb_event
        sel = tree.selection() if tree else ()
        if not sel:
            return
        for i in sorted((tree.index(s) for s in sel), reverse=True):
            del ev["attachments"][i]
        self._wb_save()
        self._wb_show_step(self._wb_step)

    def _wb_check_attachments2(self):
        missing = [a.get("path", "") for a in self._wb_event.get("attachments", [])
                   if not (a.get("path") and Path(a["path"]).exists())]
        total = len(self._wb_event.get("attachments", []))
        fp = self._fingerprint_notice()
        fp_text = ("\n\n[附件重复提示] 同一封邮件内以下文件内容完全相同"
                   "（不同路径），两份都会保留，请人工确认：\n"
                   + "\n".join(fp)) if fp else ""
        if not missing:
            (messagebox.showwarning if fp else messagebox.showinfo)(
                "附件检查", "全部就绪：%d 个附件文件均存在。" % total + fp_text)
        else:
            messagebox.showwarning("附件检查",
                                   "%d/%d 个就绪，以下不存在：\n\n%s\n\n"
                                   "缺失的附件在生成草稿时会被跳过并记入日志。"
                                   % (total - len(missing), total,
                                      "\n".join(missing[:15])) + fp_text)

    # ---------- 步骤 4：邮件模板 ----------

    def _wb_step_mail(self, parent):
        ev = self._wb_event
        mail = ev.setdefault("mail", {})
        ttk.Label(parent, text="邮件模板", style="Head.TLabel").pack(anchor="w")
        ttk.Label(parent, text="与流水线字段共用同一套字段表。",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 6))

        top = ttk.Frame(parent)
        top.pack(fill="x", pady=2)
        ttk.Label(top, text="主题：").pack(side="left")
        self.mail_subject_var = tk.StringVar(value=mail.get("subject", ""))
        ttk.Entry(top, textvariable=self.mail_subject_var, width=70).pack(side="left", padx=4)

        row2 = ttk.Frame(parent)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="代发地址：").pack(side="left")
        self.mail_from_var = tk.StringVar(value=mail.get("from_addr", ""))
        ttk.Entry(row2, textvariable=self.mail_from_var, width=34).pack(side="left", padx=4)
        ttk.Label(row2, text="默认抄送：").pack(side="left", padx=(10, 0))
        self.mail_cc_var = tk.StringVar(value=mail.get("cc_default", ""))
        ttk.Entry(row2, textvariable=self.mail_cc_var, width=26).pack(side="left", padx=4)
        ttk.Label(row2, text="默认密送：").pack(side="left", padx=(10, 0))
        self.mail_bcc_var = tk.StringVar(value=mail.get("bcc_default", ""))
        ttk.Entry(row2, textvariable=self.mail_bcc_var, width=26).pack(side="left", padx=4)

        self._build_editor_toolbar(parent)

        bot = ttk.Frame(parent)
        bot.pack(side="bottom", fill="x", pady=4)
        ttk.Button(bot, text="保存模板", command=self._wb_save_mail).pack(side="left")
        ttk.Button(bot, text="导入HTML文件…", command=self._wb_import_html).pack(side="left", padx=6)
        ttk.Button(bot, text="恢复默认", command=self._wb_reset_mail).pack(side="left")
        ttk.Button(bot, text="刷新预览", command=self._wb_refresh_preview).pack(side="left", padx=6)
        ttk.Label(bot, foreground="#888", text=(
            "可用字段：收件人名称 / 水印名称 / 打开口令 / 日期"
            " + 收件人自定义字段 + 事件常量（无权限口令字段）")).pack(side="left", padx=8)

        pw = ttk.PanedWindow(parent, orient="horizontal")
        pw.pack(fill="both", expand=True, pady=4)
        pvpane = tk.Frame(pw, width=340, background="#ffffff")
        pvpane.pack_propagate(False)
        ttk.Label(pvpane, text="预览（近似渲染，可拖中缝调整）", foreground="#666",
                  background="#ffffff").pack(anchor="w", padx=4, pady=2)
        self.preview_label = tk.Label(pvpane, anchor="nw", background="#ffffff")
        self.preview_label.pack(fill="both", expand=True)

        editpane = ttk.Frame(pw)
        t = tk.Text(editpane, undo=True, wrap="word", font=("Microsoft YaHei", 12),
                    spacing1=4, spacing3=4, padx=12, pady=8,
                    relief="solid", borderwidth=1)
        self.body_text = t
        t.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(editpane, command=t.yview)
        sb.pack(side="left", fill="y")
        t.configure(yscrollcommand=sb.set)
        t.bind("<<Paste>>", self._ed_rich_paste)
        t.bind("<KeyRelease>", self._ed_on_key)
        pw.add(editpane, weight=1)
        pw.add(pvpane, weight=0)
        pw.bind("<ButtonRelease-1>", lambda _e: self._wb_refresh_preview())

        load_html(t, mail.get("body", "") or DEFAULT_BODY)
        if not mail.get("subject"):
            self.mail_subject_var.set(DEFAULT_SUBJECT)
        self._wb_refresh_preview()

    def _wb_collect_mail(self):
        body = (text_to_html(self.body_text)
                if getattr(self, "body_text", None) is not None
                else (self._wb_event.get("mail") or {}).get("body", DEFAULT_BODY))
        return (self.mail_subject_var.get().strip(), body)

    def _wb_save_mail(self):
        mail = self._wb_event.setdefault("mail", {})
        mail["subject"], mail["body"] = self._wb_collect_mail()
        mail["from_addr"] = self.mail_from_var.get().strip()
        mail["cc_default"] = self.mail_cc_var.get().strip()
        mail["bcc_default"] = self.mail_bcc_var.get().strip()
        self._wb_save()
        self._wb_refresh_preview()
        messagebox.showinfo("模板", "已保存")

    def _wb_reset_mail(self):
        if not messagebox.askyesno("恢复默认", "确定覆盖当前主题与正文模板？"):
            return
        mail = self._wb_event.setdefault("mail", {})
        mail["subject"] = DEFAULT_SUBJECT
        mail["body"] = DEFAULT_BODY
        self.mail_subject_var.set(DEFAULT_SUBJECT)
        if getattr(self, "body_text", None) is not None:
            load_html(self.body_text, DEFAULT_BODY)
        self._wb_save()
        self._wb_refresh_preview()

    def _wb_import_html(self):
        path = filedialog.askopenfilename(title="导入正文 HTML",
                                          filetypes=[("HTML/文本", "*.html;*.htm;*.txt")])
        if not path:
            return
        try:
            raw = Path(path).read_text(encoding="utf-8")
        except UnicodeDecodeError:
            raw = Path(path).read_text(encoding="gbk", errors="ignore")
        m = re.search(r"<body[^>]*>(.*)</body>", raw, re.S | re.I)
        body = m.group(1).strip() if m else raw.strip()
        if not body:
            messagebox.showwarning("导入", "文件内容为空")
            return
        load_html(self.body_text, body)
        mail = self._wb_event.setdefault("mail", {})
        mail["body"] = text_to_html(self.body_text)
        self._wb_save()
        self._wb_refresh_preview()
        messagebox.showinfo("导入完成", "已导入并保存为当前模板。")

    def _wb_refresh_preview(self):
        try:
            import io
            from PIL import Image, ImageTk
            recips = self._wb_event.get("recipients", [])
            rec = recips[0] if recips else {
                "name": "【示例收件人名称】", "wm": "【示例水印名称】", "open_pw": "【示例口令】"}
            subject, body = self._wb_collect_mail()
            html = ('<div style="font-family:sans-serif;">'
                    '<p><b>标题：</b>%s</p><hr>%s</div>'
                    % (_htmlmod.escape(subject), body))
            png = render_preview_pixmap(html)
            img = Image.open(io.BytesIO(png))
            w = max(self.preview_label.winfo_width(), 200) - 4
            if w > 20 and img.width > w:
                img = img.resize((w, int(img.height * w / img.width)))
            self._preview_img = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self._preview_img, text="")
            self.preview_label.image = self._preview_img
        except Exception as e:
            self.preview_label.configure(
                image="", text="预览渲染失败：%s\n（可点「刷新预览」重试）" % e)

    # ---------- 步骤 5：总览（生成前核对） ----------

    def _wb_step_overview(self, parent):
        ev = self._wb_event
        ttk.Label(parent, text="总览", style="Head.TLabel").pack(anchor="w")
        ttk.Label(parent, text=(
            "生成前在这里核对：每位收件人（或每个合并组）最终会收到哪些文件、邮件长什么样。"),
            style="Sub.TLabel").pack(anchor="w", pady=(0, 6))

        errors, warnings, info = lp_v3.preflight(ev, self.store)
        head = ttk.Frame(parent)
        head.pack(fill="x", pady=(0, 4))
        for label, value in (("收件人", info["recipients"]),
                             ("发送模式", lp_v3.SEND_MODE_LABELS.get(
                                 lp_v3.send_mode_of(ev), "")),
                             ("流水线条目", info["steps"]),
                             ("待生成草稿", info["drafts"])):
            cell = ttk.Frame(head)
            cell.pack(side="left", padx=(0, 24))
            ttk.Label(cell, text=label, foreground="#666").pack(anchor="w")
            ttk.Label(cell, text=str(value), font=("", 13, "bold")).pack(anchor="w")
        if errors:
            ttk.Label(head, text="预检未通过：%d 个问题（见「预览与生成」）" % len(errors),
                      foreground="#b00").pack(side="right")
        elif warnings:
            ttk.Label(head, text="预检通过，有 %d 条提醒" % len(warnings),
                      foreground="#a60").pack(side="right")
        else:
            ttk.Label(head, text="预检通过", foreground="#0a0").pack(side="right")

        pw = ttk.PanedWindow(parent, orient="horizontal")
        pw.pack(fill="both", expand=True, pady=(4, 0))
        leftpane = ttk.Frame(pw)
        pw.add(leftpane, weight=1)
        rightpane = ttk.Frame(pw)
        pw.add(rightpane, weight=0)

        ttk.Label(leftpane, text="每位收件人 / 合并组最终会收到：",
                  style="Sub.TLabel").pack(anchor="w")
        cols = ["who", "files", "count", "pw"]
        heads = ["收件人 / 合并组", "附件包（文件名）", "件数", "打开口令"]
        tree = ttk.Treeview(leftpane, columns=cols, show="headings", height=12,
                            selectmode="browse")
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
        tree.column("who", width=150)
        tree.column("files", width=380)
        tree.column("count", width=50, stretch=False, anchor="center")
        tree.column("pw", width=110, stretch=False)
        tree.pack(fill="both", expand=True, pady=(2, 0))

        plan = lp_v3.preview_plan(ev)
        recips = dict((r.get("name", ""), r) for r in ev.get("recipients") or [])
        self._wb_ov_rows = []
        self._wb_ov_iids = {}
        for i, (key, unit, names) in enumerate(plan):
            rep = unit.get("rep") or {}
            rec = recips.get(rep.get("name", ""), {})
            who = unit.get("group") or rep.get("name", key)
            iid = "ov%d" % i
            self._wb_ov_rows.append((key, unit, names))
            self._wb_ov_iids[iid] = i
            tree.insert("", "end", iid=iid, values=(
                who, " + ".join(names) if names else "（无附件）",
                str(len(names)), rec.get("open_pw", "")))
        if not plan:
            ttk.Label(leftpane, text="还没有收件人或没有条目，先去前几步配置。",
                      style="Sub.TLabel").pack(anchor="w", pady=6)
        tree.bind("<<TreeviewSelect>>", lambda e: self._wb_overview_show(tree))
        if self._wb_ov_iids:
            tree.selection_set(sorted(self._wb_ov_iids.keys(),
                                      key=lambda x: int(x[2:]))[0])

        ttk.Label(rightpane, text="邮件预览（按左侧选中行）：",
                  style="Sub.TLabel").pack(anchor="w")
        self._wb_ov_subject = tk.Label(rightpane, anchor="w", justify="left",
                                       wraplength=440, background="#ffffff")
        self._wb_ov_subject.pack(fill="x", pady=(2, 2))
        self._wb_ov_body = tk.Label(rightpane, anchor="nw", justify="left",
                                    background="#ffffff")
        self._wb_ov_body.pack(fill="both", expand=True)
        # 完整附件清单（表格列宽有限会截断，这里给全名）
        self._wb_ov_files = tk.Text(rightpane, height=7, state="disabled",
                                    font=("Consolas", 9), background="#fafafa",
                                    wrap="word")
        self._wb_ov_files.pack(fill="x", pady=(6, 0))
        self._wb_overview_show(tree)

    def _wb_overview_show(self, tree):
        sel = tree.selection()
        if not sel:
            return
        idx = getattr(self, "_wb_ov_iids", {}).get(sel[0])
        rows = getattr(self, "_wb_ov_rows", [])
        if idx is None or idx >= len(rows):
            return
        _key, unit, names = rows[idx]
        ev = self._wb_event
        mail = ev.get("mail") or {}
        subject_tpl = mail.get("subject") or DEFAULT_SUBJECT
        body_tpl = mail.get("body") or DEFAULT_BODY
        rep = unit.get("rep") or {}
        fmap = lp_v3.build_field_map(rep, ev)
        owner_pw = rep.get("owner_pw", "")
        subject = lp_v3.render_with_fields(subject_tpl, fmap, owner_pw)
        body = lp_v3.render_with_fields(body_tpl, fmap, owner_pw, escape=True)
        self._wb_ov_subject.configure(text="主题：" + subject)
        try:
            import io
            from PIL import Image, ImageTk
            html = ('<div style="font-family:sans-serif;">%s</div>' % body)
            png = render_preview_pixmap(html)
            img = Image.open(io.BytesIO(png))
            w = max(self._wb_ov_body.winfo_width(), 220) - 4
            if w > 20 and img.width > w:
                img = img.resize((w, int(img.height * w / img.width)))
            self._wb_ov_img = ImageTk.PhotoImage(img)
            self._wb_ov_body.configure(image=self._wb_ov_img, text="")
            self._wb_ov_body.image = self._wb_ov_img
        except Exception as e:
            self._wb_ov_body.configure(image="", text=body)
        who = unit.get("group") or rep.get("name", "")
        self._wb_ov_subject.configure(
            text="【%s】主题：%s" % (who, subject))
        # 完整附件清单
        try:
            box = self._wb_ov_files
            box.configure(state="normal")
            box.delete("1.0", "end")
            if names:
                for n in names:
                    box.insert("end", "· " + n + "\n")
            else:
                box.insert("end", "（这位收件人没有附件）\n")
            box.configure(state="disabled")
        except Exception:
            pass

    # ---------- 步骤 6：预览与生成 ----------

    def _wb_step_run(self, parent):
        ev = self._wb_event
        ttk.Label(parent, text="预览与生成", style="Head.TLabel").pack(anchor="w")
        ttk.Label(parent, text="预检通过后一键生成全部 Outlook 草稿（只建草稿，不发送）。",
                  style="Sub.TLabel").pack(anchor="w", pady=(0, 6))

        # --- 预检结果 ---
        pf = ttk.LabelFrame(parent, text=" 预检结果 ", style="Card.TLabelframe")
        pf.pack(fill="x", pady=(0, 4))
        errors, warnings, info = lp_v3.preflight(ev, self.store)
        stats = ttk.Frame(pf)
        stats.pack(fill="x", padx=8, pady=4)
        for label, value in (("收件人", info["recipients"]), ("流水线条目", info["steps"]),
                              ("附件", info["attachments"]), ("待生成草稿", info["drafts"])):
            cell = ttk.Frame(stats)
            cell.pack(side="left", padx=(0, 24))
            ttk.Label(cell, text=label, foreground="#666").pack(anchor="w")
            ttk.Label(cell, text=str(value), font=("", 13, "bold")).pack(anchor="w")
        tags = ttk.Frame(pf)
        tags.pack(fill="x", padx=8, pady=(0, 6))
        if errors:
            ttk.Label(tags, text="预检未通过：%d 个问题" % len(errors),
                      foreground="#b00").pack(anchor="w")
            txt = tk.Text(pf, height=min(8, max(3, len(errors) + len(warnings))),
                          width=100, state="disabled", font=("Consolas", 9),
                          background="#fafafa")
            txt.pack(fill="x", padx=8, pady=(0, 6))
            txt.configure(state="normal")
            for e in errors:
                txt.insert("end", "[错误] " + e + "\n")
            for w in warnings:
                txt.insert("end", "[提醒] " + w + "\n")
            txt.configure(state="disabled")
        else:
            ttk.Label(tags, text="预检通过", foreground="#0a0").pack(anchor="w")
            for w in warnings[:6]:
                ttk.Label(tags, text="· " + w, foreground="#a60").pack(anchor="w")
            if len(warnings) > 6:
                ttk.Label(tags, text="…还有 %d 条提醒" % (len(warnings) - 6),
                          foreground="#a60").pack(anchor="w")

        # --- 产出清单预览 ---
        pl = ttk.LabelFrame(parent, text=" 产出清单预览 ", style="Card.TLabelframe")
        pl.pack(fill="both", expand=True, pady=(0, 4))
        cols = ["name", "files", "group", "pw", "state"]
        heads = ["收件人 / 合并组", "邮件附件包", "合并组", "打开口令", "状态"]
        tree = ttk.Treeview(pl, columns=cols, show="headings", height=8)
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
        tree.column("name", width=180)
        tree.column("files", width=380)
        tree.column("group", width=90)
        tree.column("pw", width=100)
        tree.column("state", width=90)
        tree.pack(fill="both", expand=True, padx=8, pady=4)
        plan = lp_v3.preview_plan(ev)
        recips = {r.get("name", ""): r for r in ev.get("recipients", [])}
        for key, unit, names in plan:
            rep = unit.get("rep") or {}
            rec = recips.get(rep.get("name", ""), {})
            tree.insert("", "end", values=[
                unit.get("group") or rep.get("name", key),
                " + ".join(names),
                unit.get("group", "") or "—",
                rec.get("open_pw", ""),
                "待生成"])

        # --- 执行日志 ---
        logf = ttk.LabelFrame(parent, text=" 执行日志 ", style="Card.TLabelframe")
        logf.pack(fill="both", expand=True, pady=(0, 4))
        self._wb_log_text = tk.Text(logf, height=8, state="disabled",
                                    font=("Consolas", 9), background="#fafafa")
        self._wb_log_text.pack(fill="both", expand=True, padx=8, pady=4)
        for line in (self._wb_last_run or {}).get("logs", [])[-40:]:
            self._wb_log_text.configure(state="normal")
            self._wb_log_text.insert("end", line + "\n")
            self._wb_log_text.configure(state="disabled")

    # ---------- 底栏动作 ----------

    def _wb_preflight(self):
        ev = self._wb_event
        if not ev:
            return
        errors, warnings, info = lp_v3.preflight(ev, self.store)
        if errors:
            messagebox.showerror("预检未通过",
                                 "以下问题必须先修正：\n\n"
                                 + "\n".join("· " + e for e in errors[:15])
                                 + ("\n…" if len(errors) > 15 else ""))
            return
        msg = "预检通过。\n\n收件人 %d 位 · 流水线 %d 个条目 · 附件 %d 个 · 待生成草稿 %d 封"
        msg = msg % (info["recipients"], info["steps"], info["attachments"], info["drafts"])
        if warnings:
            msg += "\n\n提醒：\n" + "\n".join("· " + w for w in warnings[:8])
        messagebox.showinfo("预检结果", msg)

    def _wb_generate(self):
        self._wb_run(None)

    def _wb_rerun_failed(self):
        run = self.store.load_last_run(self._wb_event.get("id"))
        if not run:
            messagebox.showinfo("重跑", "这个事件还没有执行记录，没有可重跑的失败项。")
            return
        names = lp_v3.failed_recipients(run)
        if not names:
            messagebox.showinfo("重跑", "上次执行没有失败项。")
            return
        if not messagebox.askyesno(
                "只重跑失败项",
                "上次执行有 %d 位收件人失败：\n\n%s\n\n只重跑这些收件人？"
                % (len(names), "\n".join(sorted(names)[:10]))):
            return
        self._wb_run(names)

    def _wb_run(self, only_recips):
        """后台线程执行事件；结果回主线程刷新。"""
        if self.running:
            messagebox.showinfo("正在执行", "上一次执行还没结束，请等它完成。")
            return
        ev = self._wb_event
        errors, _warnings, _info = lp_v3.preflight(ev, self.store)
        if errors:
            messagebox.showerror("预检未通过",
                                 "以下问题必须先修正：\n\n"
                                 + "\n".join("· " + e for e in errors[:15]))
            return
        if not messagebox.askyesno(
                "确认生成",
                "即将为「%s」生成全部文件与 Outlook 草稿。\n\n"
                "草稿只写入草稿箱，不会自动发送。\n\n开始？" % ev.get("name", "")):
            return
        self.running = True
        self._cancel = threading.Event()
        self._wb_set_running(True)
        self.wb_progress.configure(value=0)
        self.wb_status.configure(text="正在执行…")
        if getattr(self, "_wb_log_text", None) is not None:
            self._wb_log_text.configure(state="normal")
            self._wb_log_text.delete("1.0", "end")
            self._wb_log_text.configure(state="disabled")

        def job(get_word):
            run = lp_v3.run_event(self.store, ev, only_recips=only_recips,
                                  log=lambda m: self._q.put(("log", m)),
                                  progress=lambda p: self._q.put(("prog", p)),
                                  word=get_word())
            self._q.put(("run_done", (run, ev)))

        self._submit_job(job)

    # ---------------- 常驻工作线程 ----------------
    # 为什么要常驻：WordBatch 用 DispatchEx 起**私有** Word 进程，而 close() 只
    # 释放引用、不调 Quit（Quit 后再释放 COM 引用会触发 0x800706be 致命异常，
    # 直接杀进程）。所以「每次生成开一个新线程 + 新建 WordBatch」会让
    # WINWORD.EXE 越堆越多（每个约 170MB），实测能堆到 71 个把机器拖垮。
    # 改成常驻线程复用同一个 Word 实例后，整个 App 生命周期只留 1 个 Word 进程。

    def _submit_job(self, fn):
        """把任务丢给常驻工作线程（第一次调用时启动线程）。"""
        if self._worker_thread is None or not self._worker_thread.is_alive():
            self._worker_thread = threading.Thread(target=self._worker_loop,
                                                    daemon=True)
            self._worker_thread.start()
        self._jobs.put(fn)

    def _worker_loop(self):
        import pythoncom
        pythoncom.CoInitialize()
        word_box = [None]

        def get_word():
            # 在工作线程里创建（COM 对象不能跨线程用）
            if word_box[0] is None:
                word_box[0] = wm.WordBatch()
            return word_box[0]

        try:
            while True:
                fn = self._jobs.get()
                if fn is None:
                    break
                try:
                    fn(get_word)
                except Exception as e:
                    self._q.put(("log", "[致命错误] %s" % e))
                    self._q.put(("run_done", None))
                finally:
                    self._q.put(("done", None))
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # ================= 富文本编辑器（复用 v2 实现） =================

    def _build_editor_toolbar(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(4, 2))
        for label, opt in [("B", "b"), ("I", "i"), ("U", "u")]:
            tk.Button(bar, text=label, width=3,
                      font=("", 10, "bold" if opt == "b" else "normal"),
                      relief="groove",
                      command=lambda o=opt: self._ed_toggle(o)).pack(side="left", padx=1)
        ttk.Label(bar, text=" 字号").pack(side="left")
        self._ed_size = ttk.Combobox(bar, width=3, state="readonly",
                                     values=[str(x) for x in (10, 12, 14, 16, 18, 20, 24)])
        self._ed_size.set("14")
        self._ed_size.pack(side="left", padx=2)
        self._ed_size.bind("<Return>", lambda _e: self._ed_restyle(s=int(self._ed_size.get())))
        ttk.Button(bar, text="设定", width=4,
                   command=lambda: self._ed_restyle(s=int(self._ed_size.get()))).pack(side="left", padx=1)
        ttk.Button(bar, text="文字色", width=6, command=self._ed_pick_fg).pack(side="left", padx=2)
        ttk.Button(bar, text="高亮", width=5, command=self._ed_pick_bg).pack(side="left", padx=1)
        ttk.Button(bar, text="居中", width=4,
                   command=lambda: self._ed_align("center")).pack(side="left", padx=2)
        ttk.Button(bar, text="左对齐", width=6,
                   command=lambda: self._ed_align(None)).pack(side="left", padx=1)
        ttk.Button(bar, text="清除格式", width=8,
                   command=lambda: self._ed_restyle(b=None, i=None, u=None, c=None,
                                                    s=None, bg=None)).pack(side="left", padx=8)
        flds = ttk.Menubutton(bar, text="插入字段 ▾")
        m = tk.Menu(flds, tearoff=0)
        for fld in FIELDS:
            m.add_command(label=fld, command=lambda x=fld: self.body_text.insert("insert", x))
        for fld in self._wb_field_columns():
            ph = "{{%s}}" % fld
            m.add_command(label=ph, command=lambda x=ph: self.body_text.insert("insert", x))
        consts = ((self._wb_event or {}).get("options") or {}).get("constants") or {}
        for fld in consts:
            ph = "{{%s}}" % fld
            m.add_command(label=ph + "（常量）",
                          command=lambda x=ph: self.body_text.insert("insert", x))
        flds["menu"] = m
        flds.pack(side="left", padx=2)

    def _ed_need_sel(self):
        try:
            return self.body_text.index("sel.first"), self.body_text.index("sel.last")
        except tk.TclError:
            return None

    def _ed_restyle(self, **changes):
        t = self.body_text
        sel = self._ed_need_sel()
        if not sel:
            t.configure(background="#fff4f4")
            t.after(300, lambda: t.configure(background="white"))
            return
        a, b = sel
        st = _style_at(t, a)
        for k in list(_TAG_STYLE):
            t.tag_remove(k, a, b)
        st.update(changes)
        st = {k: v for k, v in st.items() if v}
        t.tag_add(_ensure_tag(t, st), a, b)

    def _ed_toggle(self, opt):
        sel = self._ed_need_sel()
        if not sel:
            return
        cur = _style_at(self.body_text, sel[0]).get(opt)
        self._ed_restyle(**{opt: None if cur else 1})

    def _ed_align(self, mode):
        """对齐作用于整行（tk 的 justify 需要覆盖整行才生效）。"""
        t = self.body_text
        sel = self._ed_need_sel()
        if not sel:
            return
        a = t.index("%s linestart" % sel[0])
        b = t.index("%s lineend" % sel[1])
        for k, st in list(_TAG_STYLE.items()):
            if st.get("al"):
                t.tag_remove(k, a, b)
        if mode:
            t.tag_add(_ensure_tag(t, {"al": mode}), a, b)

    def _ed_pick_fg(self):
        from tkinter import colorchooser
        c = colorchooser.askcolor(parent=self.root, title="文字颜色")
        if c and c[1]:
            self._ed_restyle(c=c[1])

    def _ed_pick_bg(self):
        from tkinter import colorchooser
        c = colorchooser.askcolor(parent=self.root, title="背景高亮",
                                  initialcolor="#f5f6f8")
        if c and c[1]:
            self._ed_restyle(bg=c[1])

    def _ed_rich_paste(self, _e=None):
        """富文本粘贴：优先读剪贴板 HTML 格式，保留样式。"""
        frag = clip_html_fragment()
        if not frag:
            return None  # 无 HTML 格式则走默认纯文本粘贴
        t = self.body_text
        text, runs = html_to_runs(frag)
        pos = t.index("insert")
        t.insert(pos, text)
        for a, b, st in runs:
            if a < b:
                t.tag_add(_ensure_tag(t, st), "%s+%dc" % (pos, a), "%s+%dc" % (pos, b))
        self._ed_on_key()
        return "break"

    def _ed_on_key(self, _e=None):
        if getattr(self, "_ed_after", None):
            self.root.after_cancel(self._ed_after)
        self._ed_after = self.root.after(1200, self._wb_refresh_preview)

    # ================= 水印设置 =================

    LAYOUT_OPTS = [("自动（推荐）", "auto"), ("强制单行", "single"), ("强制折行", "wrap")]
    REPEAT_OPTS = [("单个", 1), ("三份（防裁剪）", 3), ("平铺 2列×3行", "tile")]
    ANGLE_OPTS = [("自动（沿对角线）", -1.0), ("45°", 45.0), ("30°", 30.0), ("水平", 0.0)]
    FONT_OPTS = [("默认", ""), ("黑体", "黑体"), ("楷体", "楷体"),
                 ("仿宋", "仿宋"), ("宋体", "宋体"), ("微软雅黑", "微软雅黑")]

    def _update_wm_summary(self):
        if not hasattr(self, "wm_summary"):
            return
        c = dict(wm.WM_DEFAULTS, **(self.settings.get("wm_settings") or {}))
        layout = dict(self.LAYOUT_OPTS).get(c.get("layout"), c.get("layout"))
        rep = {1: "单个", 3: "三份", "tile": "平铺"}.get(c.get("repeat"), c.get("repeat"))
        a = c.get("angle")
        ang = "对角线" if a is None or float(a) < 0 else "%.0f°" % float(a)
        fs = c.get("fs_max") or "自动"
        font = c.get("font") or "默认"
        self.wm_summary.configure(
            text="%s · %s · 字号上限%s · %s · %s · 浓度%.2f"
                 % (layout, font, fs, ang, rep, float(c.get("opacity", 0.25))))

    def open_wm_settings(self):
        cur = dict(wm.WM_DEFAULTS, **(self.settings.get("wm_settings") or {}))
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

        lay = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.LAYOUT_OPTS])
        lay.set(dict((v, n) for n, v in self.LAYOUT_OPTS).get(cur.get("layout"), "自动（推荐）"))
        add_row("布局方式：", lay, hint="自动=可读则单行缩号，否则折行")
        ToolTip(lay, "自动：名称短就单行大字，太长自动折行，保证完整不截断（推荐）。\n"
                     "强制单行：无论多长都缩成一行（超长名会变得很小）。\n"
                     "强制折行：保持大字号、拆成多行。")
        self._wm_vars["layout"] = lay

        fnt = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.FONT_OPTS])
        fnt.set(dict((v, n) for n, v in self.FONT_OPTS).get(cur.get("font"), "默认"))
        add_row("字体：", fnt)
        self._wm_vars["font"] = fnt

        fsbox = ttk.Frame(frm)
        auto_fs = tk.BooleanVar(value=not cur.get("fs_max"))
        fs_spin_var = tk.IntVar(value=cur.get("fs_max") or 34)
        ttk.Checkbutton(fsbox, text="自动", variable=auto_fs).pack(side="left")
        ttk.Spinbox(fsbox, from_=16, to=48, width=4, textvariable=fs_spin_var).pack(side="left", padx=4)
        add_row("最大字号：", fsbox)
        self._wm_vars["_auto_fs"], self._wm_vars["fs_max"] = auto_fs, fs_spin_var

        ang = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.ANGLE_OPTS])
        _a = cur.get("angle")
        cur_ang = -1.0 if _a is None else float(_a)
        ang.set(dict((v, n) for n, v in self.ANGLE_OPTS).get(cur_ang, "自动（沿对角线）"))
        add_row("角度：", ang)
        ToolTip(ang, "自动：沿页面对角线（充分利用页面宽度）。\n45°/30°：更常见的固定斜角。\n"
                     "水平：不倾斜，正式感更强但遮挡正文更集中。")
        self._wm_vars["angle"] = ang

        rep = ttk.Combobox(frm, state="readonly", width=18,
                           values=[n for n, _ in self.REPEAT_OPTS])
        rep.set({1: "单个", 3: "三份（防裁剪）", "tile": "平铺 2列×3行"}
                .get(cur.get("repeat"), "单个"))
        add_row("份数：", rep)
        ToolTip(rep, "单个：页面中央一条主水印（默认）。\n\n"
                     "三份（防裁剪）：沿对角线放 3 处（中间大、两侧小更淡）。"
                     "单个水印只占中央，对方若只截图/裁剪页面的边角区域，"
                     "截图里可能一条水印都没有、无法溯源。\n\n"
                     "平铺 2列×3行：整页 6 处，防裁剪最强、存在感也最强。")
        self._wm_vars["repeat"] = rep

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
        add_row("颜色：", cbox)
        self._wm_vars["color"] = color_var

        opv = tk.DoubleVar(value=float(cur.get("opacity", 0.25)))
        op = ttk.Spinbox(frm, from_=0.05, to=0.50, increment=0.05, width=5, textvariable=opv)
        add_row("不透明度：", op)
        self._wm_vars["opacity"] = opv

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

        ttk.Separator(win).pack(fill="x", padx=10)
        pf = ttk.Frame(win)
        pf.pack(padx=10, pady=6, fill="both", expand=True)
        self._wm_preview_label = tk.Label(pf, anchor="nw", background="#ffffff")
        self._wm_preview_label.pack(side="left", fill="both", expand=True)

        def refresh_preview():
            text = "新疆乌鲁木齐市红杉泓远智能科技有限公司"
            try:
                import io
                from PIL import Image, ImageTk
                doc = fitz.open()
                page = doc.new_page()
                wm.add_watermark(page, text, collect())
                png = page.get_pixmap(dpi=72).tobytes("png")
                doc.close()
                img = Image.open(io.BytesIO(png))
                w = 340
                img = img.resize((w, int(img.height * w / img.width)))
                self._wm_preview_img = ImageTk.PhotoImage(img)
                self._wm_preview_label.configure(image=self._wm_preview_img, text="")
            except Exception as e:
                self._wm_preview_label.configure(image="", text="预览失败：%s" % e)

        ttk.Button(pf, text="刷新预览", command=refresh_preview).pack(side="left", padx=8, anchor="n")
        btns = ttk.Frame(win)
        btns.pack(pady=8)
        ttk.Button(btns, text="保存", command=lambda: save_and_close()).pack(side="left", padx=8)

        def save_and_close():
            cfg_now = collect()
            if not re.fullmatch(r"#[0-9a-fA-F]{6}", cfg_now.get("color", "")):
                messagebox.showwarning("颜色", "颜色须为 #RRGGBB 六位十六进制（如 #737373）",
                                       parent=win)
                return
            if not (0.03 <= cfg_now["opacity"] <= 0.60):
                messagebox.showwarning("不透明度", "不透明度须在 0.03 ~ 0.60 之间", parent=win)
                return
            self.settings["wm_settings"] = cfg_now
            self.store.save_settings(self.settings)
            self._update_wm_summary()
            win.destroy()

        refresh_preview()

    # ================= 更新 / 队列 / 关闭 =================

    @staticmethod
    def _ver_tuple(v):
        nums = re.findall(r"\d+", str(v or ""))
        return tuple(int(x) for x in nums[:3])

    def check_latest(self, timeout=10):
        """查询 GitHub 最新 Release。返回 dict 或 None（两种方式都失败时）。

        v3.4.7：不再只依赖 GitHub API。
        原因：匿名 API 限 60 次/小时/IP，本机走代理时出口 IP 是共享的，经常一整天
        都是 403 "API rate limit exceeded"；而 github.com 网页不受此限。
        因此先请求网页 /releases/latest 跟随 302 拿 tag，API 只用来补充说明与资产摘要，
        403/超时时静默降级（版本号仍然可用）。
        """
        import urllib.request

        self._last_update_err = ""
        tag = ""
        html_url = "https://github.com/%s/releases" % UPDATE_REPO

        # --- 1) 网页跳转：不受 API 限流（浏览器能开 GitHub，这里就能通） ---
        try:
            web = "https://github.com/%s/releases/latest" % UPDATE_REPO
            req = urllib.request.Request(web, headers={"User-Agent": "doc-mailer"})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                final = r.geturl()
                m = re.search(r"/releases/tag/([^/?#]+)", final or "")
                if m:
                    tag = m.group(1)
                    html_url = final
        except Exception as e:
            self._last_update_err = "网页跳转：%s" % e

        # --- 2) API：补充更新说明 / 资产大小与 digest（拿不到不算失败） ---
        body, asset_url, asset_size, asset_digest = "", None, None, None
        try:
            api = "https://api.github.com/repos/%s/releases/latest" % UPDATE_REPO
            headers = {"User-Agent": "doc-mailer",
                       "Accept": "application/vnd.github+json"}
            tok = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
            if tok:      # 有 token 时用认证额度（5000 次/小时），没有也不影响版本检测
                headers["Authorization"] = "Bearer " + tok
            req = urllib.request.Request(api, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                d = json.loads(r.read().decode("utf-8"))
            tag = tag or str(d.get("tag_name") or "")
            body = (d.get("body") or "")[:400]
            html_url = d.get("html_url") or html_url
            asset = next((a for a in d.get("assets", [])
                          if str(a.get("name", "")).startswith("doc-mailer_")
                          and str(a.get("name", "")).endswith("_portable.zip")), None)
            if asset:
                asset_url = asset.get("browser_download_url")
                asset_size = asset.get("size")
                asset_digest = asset.get("digest")
        except Exception as e:
            msg = "GitHub API：%s" % e
            self._last_update_err = ((self._last_update_err + "；")
                                     if self._last_update_err else "") + msg

        if not tag:
            return None
        if not html_url or "/releases/tag/" not in html_url:
            html_url = "https://github.com/%s/releases/tag/%s" % (UPDATE_REPO, tag)
        ver = tag[1:] if tag[:1] in ("v", "V") else tag
        if not asset_url and re.match(r"^\d+(\.\d+)*$", ver):
            # 资产命名规范固定（HANDOFF §7），API 拿不到时按规范拼下载地址
            asset_url = ("https://github.com/%s/releases/download/%s/"
                         "doc-mailer_v%s_portable.zip" % (UPDATE_REPO, tag, ver))
        return {"tag": tag, "body": body, "url": html_url,
                "asset_url": asset_url, "asset_size": asset_size,
                "asset_digest": asset_digest}

    def check_update(self, auto=False):
        def work():
            try:
                info = self.check_latest()
            except Exception as e:
                info = None
                self._last_update_err = "未知异常：%s" % e
            if info:
                self._q.put(("update", (info, auto)))
            elif not auto:
                detail = str(getattr(self, "_last_update_err", "") or "无详细信息")
                if len(detail) > 300:
                    detail = detail[:300] + "…"
                self._q.put(("update_err",
                             "无法获取 GitHub 最新版本。\n\n"
                             "已尝试「网页跳转 + GitHub API」两种方式。\n"
                             "明细：%s\n\n"
                             "可直接打开 Releases 页面手动下载。" % detail))
        threading.Thread(target=work, daemon=True).start()

    def _show_update(self, info, auto):
        remote = self._ver_tuple(info["tag"])
        local = self._ver_tuple(APP_VERSION)
        if remote <= local:
            if not auto:
                messagebox.showinfo("检查更新", "当前已是最新版本（v%s）。" % APP_VERSION)
            return
        frozen = getattr(sys, "frozen", False)
        msg = ("发现新版本：%s（当前 v%s）\n\n" % (info["tag"], APP_VERSION)
               + info["body"].replace("\r", "")[:300] + "\n")
        if frozen or not info.get("asset_url"):
            msg += "\n点击「打开下载页」获取最新版本。"
            if messagebox.askyesno("发现新版本", msg + "\n\n现在打开下载页？"):
                webbrowser.open(info["url"])
            return
        msg += "\n\n「自动更新」将下载新版并自动替换（保留你的数据与清单），完成后自动重启。"
        if not messagebox.askyesno("发现新版本", msg + "\n\n现在自动更新？"):
            return
        self._self_update(info)

    def _self_update(self, info):
        def work():
            import zipfile
            import hashlib
            staging = app_dir() / ".update_staging"
            zip_path = app_dir() / ".update_pkg.zip"
            try:
                self._q.put(("log", "[更新] 下载 %s（约 %dMB）…"
                             % (info["tag"], (info.get("asset_size") or 0) // 1048576)))
                last = [None]

                def _prog(pct):
                    if pct != last[0] and pct % 20 == 0:
                        last[0] = pct
                        self._q.put(("log", "[更新] 下载进度 %d%%" % pct))

                import urllib.request
                urllib.request.urlretrieve(info["asset_url"], zip_path, reporthook=_prog)
                raw = zip_path.read_bytes()
                digest = (info.get("asset_digest") or "")
                if digest.startswith("sha256:"):
                    got = "sha256:" + hashlib.sha256(raw).hexdigest()
                    if got != digest:
                        raise RuntimeError("下载校验失败（期望 %s…，实际 %s…）"
                                           % (digest[:20], got[:20]))
                    self._q.put(("log", "[更新] SHA256 校验通过"))
                if staging.exists():
                    import shutil as _sh
                    _sh.rmtree(staging)
                with zipfile.ZipFile(zip_path) as z:
                    z.extractall(staging)
                src_dir = next((d for d in staging.iterdir() if d.is_dir()), staging)
                # 更新脚本：等待程序退出 → 覆盖（保留用户数据）→ 重启
                # 排除清单必须含 lp_data（v3 的用户数据目录）与旧版单文件，
                # 否则 robocopy /MOVE 会把用户数据一起覆盖掉。
                # 只能有**一条** robocopy：先复制再排除就已经晚了。
                bat = app_dir() / "run_update.bat"
                NLc = chr(10)
                bat.write_text(
                    "@echo off" + NLc +
                    "chcp 65001 >nul 2>&1" + NLc +
                    "cd /d \"%~dp0\"" + NLc +
                    "ping -n 4 127.0.0.1 >nul" + NLc +
                    'robocopy ".update_staging' + chr(92) + str(src_dir.name) + '" "." /E /MOVE '
                    "/XF lp_data.json lp_run_log.csv lp_data.json.bak "
                    "lp_data.json.tmp lp_data.example.json "
                    "/XD lp_data >nul" + NLc +
                    'rd /s /q ".update_staging" 2>nul' + NLc +
                    'del ".update_pkg.zip" 2>nul' + NLc +
                    'start "" "%~dp0start.bat"' + NLc +
                    "exit" + NLc, encoding="ascii", errors="ignore")
                zip_path.unlink()
                self._q.put(("update_ready", None))
            except Exception as e:
                try:
                    zip_path.unlink(missing_ok=True)
                except OSError:
                    pass
                self._q.put(("update_err",
                             "自动更新失败：%s\n可手动到 Releases 页下载。" % e))
        threading.Thread(target=work, daemon=True).start()

    def log(self, msg):
        # 后台线程只投递消息，主线程轮询写入（Tk 非线程安全）
        self._q.put(("log", msg))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "log":
                    self._wb_log(payload)
                elif kind == "prog":
                    try:
                        self.wb_progress.configure(value=payload)
                    except Exception:
                        pass
                    bar2 = getattr(self, "_wb_progress2", None)
                    if bar2 is not None:
                        try:
                            bar2.configure(value=payload)
                        except Exception:
                            pass
                elif kind == "run_done":
                    if isinstance(payload, tuple):
                        run_obj, run_ev = payload
                    else:
                        run_obj, run_ev = payload, self._wb_event
                    self._wb_last_run = run_obj
                    if run_obj is not None:
                        try:
                            self._apply_process_results(run_obj, run_ev)
                        except Exception:
                            pass
                    if run_obj is not None:
                        st = run_obj.get("stats") or {}
                        self.wb_status.configure(
                            text="成功 %d · 失败 %d · 跳过 %d"
                                 % (st.get("ok", 0), st.get("fail", 0), st.get("skip", 0)))
                        try:
                            # 回到「预览与生成」页
                            self._wb_show_step(WB_STEP_RUN)
                        except Exception:
                            pass
                elif kind == "summary":
                    self._wb_log(payload)
                    self.root.after(300, lambda p=payload: messagebox.showinfo("批量结束", p))
                elif kind == "update":
                    try:
                        self._show_update(*payload)
                    except Exception as e:
                        # v3.4.7：更新信息处理异常不能静默（之前被外层 except 吞掉）
                        messagebox.showerror("检查更新",
                                             "处理更新信息失败：%s" % e)
                elif kind == "update_err":
                    messagebox.showwarning("检查更新", payload)
                elif kind == "update_ready":
                    if messagebox.askyesno(
                            "更新就绪",
                            "新版已下载并通过校验。\n\n"
                            "现在退出并完成更新？（数据与清单保留，完成后自动重启）"):
                        os.startfile(str(app_dir() / "run_update.bat"))
                        self.root.destroy()
                elif kind == "done":
                    self.running = False
                    self._wb_set_running(False)
        except Exception:
            pass
        self._poll_id = self.root.after(120, self._poll_queue)

    def shutdown(self):
        """停掉后台轮询并拆掉外壳（测试与异常退出用）。"""
        if self._poll_id is not None:
            try:
                self.root.after_cancel(self._poll_id)
            except Exception:
                pass
            self._poll_id = None
        # 让常驻工作线程收工（它持有的 Word 引用也随之释放）
        try:
            self._jobs.put_nowait(None)
        except Exception:
            pass
        for w in self._shell_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        self._shell_widgets = []

    def _wb_log(self, msg):
        """写日志到当前步骤的日志框（没有就写状态栏）。"""
        box = getattr(self, "_wb_log_text", None)
        if box is not None:
            box.configure(state="normal")
            box.insert("end", msg + "\n")
            box.see("end")
            box.configure(state="disabled")
        else:
            self.wb_status.configure(text=msg[:60])

    def _wb_set_running(self, running):
        self.running = running
        btns = [self.btn_preflight, self.btn_generate, self.btn_rerun]
        pb = getattr(self, "_wb_process_btn", None)
        if pb is not None:
            btns.append(pb)
        for b in btns:
            try:
                b.state(["disabled"] if running else ["!disabled"])
            except tk.TclError:
                pass
        if not running:
            self._wb_render_mapping_table()

    def on_close(self):
        if self.running and not messagebox.askyesno(
                "批量进行中",
                "后台仍在生成文件/草稿。\n\n确定退出？（强退可能留下不完整文件，"
                "已完成的草稿保留）"):
            return
        try:
            if getattr(self, "body_text", None) is not None and self._wb_event:
                mail = self._wb_event.setdefault("mail", {})
                mail["subject"], mail["body"] = self._wb_collect_mail()
                self.store.save_event(self._wb_event)
        except Exception:
            pass
        try:
            self.store.save_settings(self.settings)
        except Exception:
            pass
        self.root.destroy()


def _install_excepthook():
    """无控制台形态下未捕获异常不再'点了没反应'：落盘 + 弹窗。"""
    import traceback

    def hook(typ, val, tb):
        try:
            logdir = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "doc-mailer"
            logdir.mkdir(parents=True, exist_ok=True)
            (logdir / "error.log").write_text(
                traceback.format_exception(typ, val, tb), encoding="utf-8")
        except Exception:
            pass
        try:
            import tkinter.messagebox as _mb
            _mb.showerror("程序出错",
                          "%s: %s\n\n详情已写入 %s" % (typ.__name__, val, logdir / "error.log"))
        except Exception:
            pass
    sys.excepthook = hook
    threading.excepthook = lambda a: hook(a.exc_type, a.exc_value, a.exc_traceback)


def main():
    _install_excepthook()
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
