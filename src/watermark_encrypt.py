#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 水印 + 双口令加密 + Word 通知模板 → PDF
"""

import argparse
import glob
import math
import os
import re
import secrets
import sys
from pathlib import Path

import fitz  # PyMuPDF

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ---------------- 水印排版 ----------------

FONT_FILES = {
    "黑体": "simhei.ttf",
    "楷体": "simkai.ttf",
    "仿宋": "simfang.ttf",
    "宋体": "simsun.ttc",
    "微软雅黑": "msyh.ttc",
}
_font_cache = {}

WM_DEFAULTS = {
    "layout": "auto",
    "fs_max": 0,
    "font": "",
    "color": "#737373",
    "opacity": 0.25,
    "angle": -1,
    "repeat": 1,
}


def _load_wm_font(name):
    if not name or name not in FONT_FILES:
        return fitz.Font("china-s")
    if name not in _font_cache:
        try:
            _font_cache[name] = fitz.Font(fontfile=str(
                Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / FONT_FILES[name]))
        except Exception:
            _font_cache[name] = fitz.Font("china-s")
    return _font_cache[name]


def _hex_to_rgb(h):
    h = (h or "#737373").lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _wrap_by_width(text, font, fontsize, max_w):
    lines, cur = [], ""
    for ch in text:
        if font.text_length(cur + ch, fontsize=fontsize) <= max_w:
            cur += ch
        else:
            if cur:
                lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines


def plan_layout(text, font, usable_w, usable_h, fs_max, layout="auto"):
    if layout == "single":
        fs = fs_max
        while fs >= 8 and font.text_length(text, fontsize=fs) > usable_w:
            fs -= 0.5
        return [text], max(fs, 8)
    if layout == "wrap":
        fs = fs_max
        while fs >= 10:
            lines = _wrap_by_width(text, font, fs, usable_w)
            if len(lines) <= 4 and len(lines) * fs * 1.4 <= usable_h:
                return lines, fs
            fs -= 1
        return _wrap_by_width(text, font, 10, usable_w), 10
    fs = fs_max
    while fs >= 10:
        if font.text_length(text, fontsize=fs) <= usable_w:
            break
        fs -= 0.5
    if fs >= max(14, fs_max * 0.55):
        return [text], fs
    fs = fs_max
    while fs >= 10:
        lines = _wrap_by_width(text, font, fs, usable_w)
        if len(lines) <= 3 and len(lines) * fs * 1.4 <= usable_h:
            return lines, fs
        fs -= 1
    return _wrap_by_width(text, font, 10, usable_w), 10


def _make_stamp(text, font, side, cfg):
    tmp = fitz.open()
    tp = tmp.new_page(width=side, height=side)
    h_limit = float(cfg.get("_h_limit") or 0)
    color = _hex_to_rgb(cfg.get("color"))
    opacity = float(cfg.get("opacity", 0.25))
    layout = cfg.get("layout", "auto")
    auto_fs = max(16, min(40, side * 0.033))
    fs_max = int(cfg.get("fs_max") or 0)
    fs_main = min(fs_max, 60) if fs_max else auto_fs

    def write_block(cx, cy, box_w, box_h, fs_cap, op):
        lines, fs = plan_layout(text, font, box_w, box_h, fs_cap, layout)
        tw = fitz.TextWriter(tp.rect)
        line_h = fs * 1.35
        n = len(lines)
        try:
            ang = float(cfg.get("angle", -1))
        except (TypeError, ValueError):
            ang = -1
        for i, ln in enumerate(lines):
            w = font.text_length(ln, fontsize=fs)
            off = (i - (n - 1) / 2) * line_h
            y = (cy + off if ang != 0 else cy - off) + fs * 0.36
            tw.append(fitz.Point(cx - w / 2, y), ln, font=font, fontsize=fs)
        tw.write_text(tp, color=color, opacity=op)

    rep = cfg.get("repeat", 1)
    if rep == "tile":
        rep = 1
    box_w = min(side * 0.82, h_limit) if h_limit else side * 0.82
    write_block(side / 2, side / 2, box_w, side * 0.70, fs_main, opacity)
    if rep == 3:
        for off in (-0.20, 0.20):
            write_block(side / 2 + off * side * 0.82, side / 2,
                        side * 0.50, side * 0.40, fs_main * 0.62, opacity * 0.6)
    return tmp


def add_watermark(page, text, cfg=None):
    cfg = dict(WM_DEFAULTS, **(cfg or {}))
    r = page.rect
    font = _load_wm_font(cfg.get("font"))
    try:
        angle = float(cfg.get("angle", -1))
    except (TypeError, ValueError):
        angle = -1
    if angle < 0:
        angle = math.degrees(math.atan2(r.height, r.width))

    if abs(angle) < 1e-6:
        cfg = dict(cfg, _h_limit=r.width * 0.92)

    def place(stamp, cx, cy, side):
        page.show_pdf_page(fitz.Rect(cx - side / 2, cy - side / 2,
                                     cx + side / 2, cy + side / 2),
                           stamp, 0, rotate=angle, overlay=True)

    if cfg.get("repeat") == "tile":
        cols, rows = 2, 3
        cw, ch = r.width / cols, r.height / rows
        cell_side = math.ceil(math.hypot(cw, ch) * 1.02)
        cell_cfg = dict(cfg, repeat=1)
        cell_cfg["fs_max"] = min(int(cfg.get("fs_max") or 40), int(ch * 0.35))
        for rr in range(rows):
            for cc in range(cols):
                stamp = _make_stamp(text, font, cell_side, cell_cfg)
                place(stamp, r.x0 + cw * (cc + 0.5), r.y0 + ch * (rr + 0.5), cell_side)
        return

    side = math.ceil(math.hypot(r.width, r.height) * 1.05)
    stamp = _make_stamp(text, font, side, cfg)
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    place(stamp, cx, cy, side)


# ---------------- 加密 ----------------

def _perm_summary(doc):
    p = doc.permissions & 0xFFFFFFFF
    return {
        "打印": bool(p & fitz.PDF_PERM_PRINT),
        "修改": bool(p & fitz.PDF_PERM_MODIFY),
        "复制文字": bool(p & fitz.PDF_PERM_COPY),
        "批注": bool(p & fitz.PDF_PERM_ANNOTATE),
        "填表单": bool(p & fitz.PDF_PERM_FORM),
        "页面拼装": bool(p & fitz.PDF_PERM_ASSEMBLE),
    }


def process_one(src: Path, out: Path, name, open_pw, owner_pw, no_encrypt,
                repeat=1, opacity=0.25, wm_cfg=None):
    cfg = dict(WM_DEFAULTS)
    if repeat in (1, 3, "tile"):
        cfg["repeat"] = repeat
    if opacity:
        cfg["opacity"] = opacity
    if wm_cfg:
        cfg.update(wm_cfg)
    if not (name or "").strip():
        raise ValueError("水印文字为空，拒绝生成无水印文件")
    if not open_pw or not owner_pw:
        raise ValueError("打开口令/权限口令不能为空")
    if open_pw == owner_pw:
        raise ValueError("权限口令不能与打开口令相同（否则收件人获得全部权限）")

    warnings = []
    with fitz.open(src) as doc:
        for page in doc:
            add_watermark(page, name, cfg)
        try:
            doc.subset_fonts()
        except Exception as e:
            warnings.append(f"字体子集化失败（体积可能偏大）：{e}")
        tmp_out = out.with_suffix(out.suffix + ".tmp")
        if no_encrypt:
            doc.save(tmp_out, garbage=3, deflate=True)
        else:
            perms = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
            if hasattr(fitz, "PDF_PERM_ACCESSIBILITY"):
                perms |= fitz.PDF_PERM_ACCESSIBILITY
            doc.save(tmp_out, encryption=fitz.PDF_ENCRYPT_AES_256,
                     user_pw=open_pw, owner_pw=owner_pw,
                     permissions=perms, garbage=3, deflate=True)
    os.replace(tmp_out, out)

    if no_encrypt:
        return {"encrypted": False, "warnings": warnings}

    d = fitz.open(out)
    try:
        if not d.needs_pass:
            raise RuntimeError("输出文件未加密！")
        if d.authenticate(open_pw) < 2:
            raise RuntimeError("打开口令验证失败")
        wm_text_found = all(pg.get_text().find(name[:6]) >= 0 for pg in d)
        perms = _perm_summary(d)
    finally:
        d.close()
    if not wm_text_found:
        raise RuntimeError("自检未检出水印文字")
    d2 = fitz.open(out)
    try:
        owner_ok = d2.authenticate(owner_pw) >= 4
    finally:
        d2.close()
    if not owner_ok:
        raise RuntimeError("权限口令验证失败")
    return {"encrypted": True, "user_perms": perms, "warnings": warnings}


# ---------------- Word 通知模板 → PDF ----------------

def generate_notice_pdf(template_path, output_path, replacements):
    """从 Word 模板生成单份 PDF（占位符替换后导出）。

    Args:
        template_path: .docx 模板完整路径
        output_path: 输出 PDF 完整路径
        replacements: dict，如 {"{{LP_Name}}": "XX基金", "{{CC1}}": "$200,000"}
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(Path(template_path).resolve()), ReadOnly=True)

        for find_text, replace_text in replacements.items():
            # 必须用位置参数：pywin32 动态调度会静默忽略 Replace=2 这种命名参数
            doc.Content.Find.Execute(
                find_text, False, False, False, False, False,
                True, 1, False, str(replace_text), 2)   # Replace=2 → wdReplaceAll

        # SaveAs2 FileFormat=17 (wdFormatPDF)
        doc.SaveAs2(str(Path(output_path).resolve()), FileFormat=17)
    except Exception as e:
        raise RuntimeError(f"Word 模板处理失败：{e}")
    finally:
        if doc:
            try:
                doc.Close(0)
            except Exception:
                pass
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def generate_notice_pdfs(template_path, output_dir, lp_records):
    """从 Word 模板批量生成 PDF（逐 LP 替换占位符）。

    Args:
        template_path: .docx 模板路径
        output_dir: 输出目录
        lp_records: LP 记录列表（lp_tool.App.data["lps"] 格式）

    Returns:
        list of {"name": str, "pdf": str or None, "error": str or None}
    """
    results = []
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    for rec in lp_records:
        name = rec.get("name", "")
        try:
            # 构造替换映射：{{name}} 和 ?LP_Name? 两种占位符都支持
            replacements = {
                "{{name}}": name,
                "{{收件人名称}}": name,
                "{{LP名称}}": name,
                "?LP_Name?": name,
                "?LP_Name_?": name,
            }
            for k, v in rec.items():
                if k in ("name",):  # 已在上面处理
                    continue
                if isinstance(v, str) and v:
                    ph = "{{" + k + "}}"
                    replacements[ph] = v
            # 数字格式化
            for num_key in ("call_amount", "unfunded"):
                val = rec.get(num_key)
                if val:
                    try:
                        replacements["{{" + num_key + "}}"] = f"{float(val):,.0f}"
                    except (ValueError, TypeError):
                        replacements["{{" + num_key + "}}"] = str(val)

            safe = _hex_to_rgb  # noqa: F841 — 防止导入被 tree-shake
            safe_name = _wrap_by_width  # noqa: F841
            fname = re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "unnamed"
            pdf_path = out_dir / (fname + ".pdf")
            generate_notice_pdf(template_path, str(pdf_path), replacements)
            results.append({"name": name, "pdf": str(pdf_path), "error": None})
        except Exception as e:
            results.append({"name": name, "pdf": None, "error": str(e)})

    return results


import re  # generate_notice_pdfs 里用到


# ---------------- 流水线用的分步接口 ----------------
#
# 把「加水印」与「加密」从 process_one 里拆成两个独立函数，
# 行为与原来完全一致（同样的水印算法、同样的权限位、同样的子集化），
# 只是可以分步调用、可以只水印不加密、或只加密不水印。
# v3.1 起上层用的是「产出条目」模型（generate / attach），
# 这里仍是底层能力（条目内部的处理链按 转 PDF → 水印 → 加密 调用它们）。


def _wm_config(repeat=1, opacity=0.25, wm_cfg=None):
    cfg = dict(WM_DEFAULTS)
    if repeat in (1, 3, "tile"):
        cfg["repeat"] = repeat
    if opacity:
        cfg["opacity"] = opacity
    if wm_cfg:
        cfg.update(wm_cfg)
    return cfg


def watermark_pdf(src, out, text, wm_cfg=None):
    """只加水印（不加密）。返回 {"warnings": [...]}。

    水印旋转走「方形临时页 + show_pdf_page(rotate=)」，
    绝不用 TextWriter.write_text(morph=...)（PyMuPDF 1.28 会静默丢字）。
    """
    cfg = dict(WM_DEFAULTS, **(wm_cfg or {}))
    if not (text or "").strip():
        raise ValueError("水印文字为空，拒绝生成无水印文件")
    warnings = []
    out = Path(out)
    with fitz.open(src) as doc:
        for page in doc:
            add_watermark(page, text, cfg)
        try:
            doc.subset_fonts()
        except Exception as e:
            warnings.append(f"字体子集化失败（体积可能偏大）：{e}")
        tmp_out = out.with_suffix(out.suffix + ".tmp")
        doc.save(tmp_out, garbage=3, deflate=True)
    os.replace(tmp_out, out)
    return {"warnings": warnings}


def permission_mask(permissions=None):
    """权限勾选 → PyMuPDF 权限位掩码。

    permissions: None 或空 → 默认「仅打印」（与既有铁律一致）。
      "print" 允许打印
      "edit"  允许编辑（修改内容）
      "copy"  允许复制内容
    """
    if permissions is None:
        permissions = ["print"]
    if isinstance(permissions, str):
        permissions = [permissions]
    perms = list(permissions)
    mask = 0
    if "print" in perms:
        mask |= fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
    if "edit" in perms:
        mask |= fitz.PDF_PERM_MODIFY
    if "copy" in perms:
        mask |= fitz.PDF_PERM_COPY | fitz.PDF_PERM_ANNOTATE
    if hasattr(fitz, "PDF_PERM_ACCESSIBILITY"):
        mask |= fitz.PDF_PERM_ACCESSIBILITY
    return mask


def encrypt_pdf(src, out, open_pw, owner_pw, permissions=None):
    """只加密（AES-256）。返回加密自检结果。

    permissions: 可配置权限位，默认 ["print"]（仅打印）。
      print=允许打印  edit=允许编辑  copy=允许复制内容
    """
    if not open_pw or not owner_pw:
        raise ValueError("打开口令/权限口令不能为空")
    if open_pw == owner_pw:
        raise ValueError("权限口令不能与打开口令相同（否则收件人获得全部权限）")
    perms = permission_mask(permissions)
    warnings = []
    out = Path(out)
    with fitz.open(src) as doc:
        try:
            doc.subset_fonts()   # 保存前子集化，否则字体全量嵌入（0.1MB → 21MB）
        except Exception as e:
            warnings.append(f"字体子集化失败（体积可能偏大）：{e}")
        tmp_out = out.with_suffix(out.suffix + ".tmp")
        doc.save(tmp_out, encryption=fitz.PDF_ENCRYPT_AES_256,
                 user_pw=open_pw, owner_pw=owner_pw,
                 permissions=perms, garbage=3, deflate=True)
    os.replace(tmp_out, out)

    d = fitz.open(out)
    try:
        if not d.needs_pass:
            raise RuntimeError("输出文件未加密！")
        if d.authenticate(open_pw) < 2:
            raise RuntimeError("打开口令验证失败")
        perms_summary = _perm_summary(d)
    finally:
        d.close()
    d2 = fitz.open(out)
    try:
        owner_ok = d2.authenticate(owner_pw) >= 4
    finally:
        d2.close()
    if not owner_ok:
        raise RuntimeError("权限口令验证失败")
    return {"encrypted": True, "user_perms": perms_summary, "warnings": warnings}


def detect_placeholders(docx_path):
    """从 .docx 里提取占位符（{{字段}} 与 ?字段? 两种写法）。

    直接读 word/document.xml，不依赖 Word：更快，也不会弹出窗口。
    Word 常把一个占位符拆进多个 <w:t>，这里把文本节点原样拼回连续文本。
    """
    import zipfile
    text = ""
    try:
        with zipfile.ZipFile(docx_path) as z:
            names = [n for n in z.namelist()
                     if n.startswith("word/") and n.endswith(".xml")
                     and ("document" in n or "header" in n or "footer" in n
                          or "footnotes" in n or "endnotes" in n)]
            names.sort(key=lambda n: (0 if n == "word/document.xml" else 1, n))
            for n in names:
                raw = z.read(n).decode("utf-8", errors="replace")
                parts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", raw, re.S)
                text += "".join(parts)
    except Exception as e:
        raise RuntimeError(f"读取 Word 模板失败：{e}")
    for a, b in (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'),
                 ("&apos;", "'"), ("&amp;", "&")):
        text = text.replace(a, b)
    found = set(re.findall(r"\{\{[^}]+\}\}", text))
    found |= set(re.findall(r"\?[^?\s]{2,30}\?", text))
    return sorted(found)


class WordBatch:
    """批量处理时复用一个 Word 实例（每份文件启一次 Word 非常慢）。

    两个必须知道的坑：
    1) 用 DispatchEx 起**私有实例**：普通 Dispatch 会挂到用户已经开着的 Word 上，
       收尾时会把用户的 Word 一起关掉。
    2) 收尾**不能调 Word.Quit()**：Quit 之后 Word 的 COM 服务器就没了，
       这时 Python 释放最后一个引用会收到 RPC_S_CALL_FAILED(0x800706be)，
       那是 Windows 致命异常，会直接把整个进程干掉（应用会在跑一半时闪退）。
       所以这里只释放引用，不 Quit——Word 实例随进程一起退出。
    """

    def __init__(self):
        self._word = None
        self._com = False

    def _ensure(self):
        if self._word is None:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()   # 新线程必须初始化 COM
            self._com = True
            # DispatchEx：私有实例，绝不碰用户自己开着的 Word
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
        return self._word

    def replace_and_save(self, template_path, out_path, replacements, fmt=16):
        """按映射替换占位符后另存。fmt: 16=docx, 17=pdf。

        替换用 Find.Execute 的**位置参数**形式（Replace=wdReplaceAll），
        只改文字不重建文档，字体/字号/颜色/表格/页眉页脚/分页全部原样保留。

        注意：不能用 find.Execute(Replace=2) 这种命名参数写法——pywin32 的动态
        调度会静默忽略命名参数，结果是「一个字都没替换」还不报错。
        """
        word = self._ensure()
        doc = word.Documents.Open(str(Path(template_path).resolve()), ReadOnly=True)
        try:
            for find_text, replace_text in (replacements or {}).items():
                # Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards,
                #   MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format,
                #   ReplaceWith, Replace, ...)
                doc.Content.Find.Execute(
                    find_text, False, False, False, False, False,
                    True, 1, False, str(replace_text), 2)
            doc.SaveAs2(str(Path(out_path).resolve()), FileFormat=fmt)
        finally:
            try:
                doc.Close(0)
            except Exception:
                pass

    def convert(self, src_docx, out_pdf):
        """docx → pdf（Word 导出，格式保持由 Word 保证）。"""
        word = self._ensure()
        doc = word.Documents.Open(str(Path(src_docx).resolve()), ReadOnly=True)
        try:
            doc.SaveAs2(str(Path(out_pdf).resolve()), FileFormat=17)
        finally:
            try:
                doc.Close(0)
            except Exception:
                pass

    def create_docx(self, path, text):
        """新建一个 .docx（示例模板生成用）。"""
        word = self._ensure()
        doc = word.Documents.Add()
        try:
            doc.Content.Text = text
            doc.SaveAs2(str(Path(path).resolve()), FileFormat=16)
        finally:
            try:
                doc.Close(0)
            except Exception:
                pass

    def close(self):
        """释放 Word 引用（**不调 Quit**，理由见类注释）。

        释放发生在 Word 还活着的时候，所以不会触发 RPC 致命异常；
        残留的 Word 实例会在本进程结束时一并退出。
        """
        word = self._word
        self._word = None
        if word is not None:
            del word
            import gc
            gc.collect()
        if self._com:
            import pythoncom
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com = False

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        self.close()
        return False


def generate_notice_docx(template_path, output_path, replacements, batch=None):
    """从 Word 模板生成 .docx（占位符替换后，格式保持不变）。"""
    own = batch is None
    b = batch or WordBatch()
    try:
        b.replace_and_save(template_path, output_path, replacements, fmt=16)
    except Exception as e:
        raise RuntimeError(f"Word 模板处理失败：{e}")
    finally:
        if own:
            b.close()


def docx_to_pdf(src_docx, output_path, batch=None):
    """docx → pdf（Word COM 导出）。"""
    own = batch is None
    b = batch or WordBatch()
    try:
        b.convert(src_docx, output_path)
    except Exception as e:
        raise RuntimeError(f"Word 转 PDF 失败：{e}")
    finally:
        if own:
            b.close()


# ---------------- CLI ----------------

def main():
    ap = argparse.ArgumentParser(description="PDF 水印 + 双口令加密")
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--name", required=True)
    ap.add_argument("--open-pw", required=True)
    ap.add_argument("--owner-pw", default=None)
    ap.add_argument("--outdir", default="out_wm")
    ap.add_argument("--repeat", default="1", choices=["1", "3", "tile"])
    ap.add_argument("--opacity", type=float, default=0.25)
    ap.add_argument("--layout", default="auto", choices=["auto", "single", "wrap"])
    ap.add_argument("--font", default="", choices=[""] + list(FONT_FILES))
    ap.add_argument("--color", default="#737373")
    ap.add_argument("--angle", type=float, default=-1)
    ap.add_argument("--fs-max", type=int, default=0)
    ap.add_argument("--no-encrypt", action="store_true")
    args = ap.parse_args()

    if args.owner_pw is None:
        owner_pw = secrets.token_urlsafe(12)
        generated = True
    else:
        owner_pw = args.owner_pw
        generated = False
    if owner_pw == args.open_pw:
        sys.exit("错误：权限口令不能与打开口令相同")

    files = []
    for pattern in args.inputs:
        files.extend(glob.glob(pattern))
    files = sorted({f for f in files if f.lower().endswith(".pdf")})
    if not files:
        sys.exit("没有找到输入 PDF")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"水印文字：{args.name}")
    for f in files:
        src = Path(f)
        out = outdir / f"{src.stem}_wm.pdf"
        info = process_one(src, out, args.name, args.open_pw, owner_pw,
                           args.no_encrypt,
                           repeat=int(args.repeat) if args.repeat != "tile" else "tile",
                           opacity=args.opacity,
                           wm_cfg={"layout": args.layout, "font": args.font,
                                   "color": args.color, "angle": args.angle,
                                   "fs_max": args.fs_max})
        if info.get("encrypted"):
            print(f"[OK] {src.name} -> {out}")
            print(f"     LP 侧权限：{info['user_perms']}")
        else:
            print(f"[OK] {src.name} -> {out}  (未加密，仅水印)")

    if not args.no_encrypt:
        print("-" * 60)
        print(f"打开口令（发 LP）：{args.open_pw}")
        print(f"权限口令（勿发 LP）{'（随机生成，请自行记录）' if generated else ''}：{owner_pw}")
        print("权限设置：允许打印；禁止修改、批注、复制文字、页面拼装")


if __name__ == "__main__":
    main()

