#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 水印 + 双口令加密 + Word 通知模板 → PDF
"""

import argparse
import glob
import math
import os
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
            find = doc.Content.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = find_text
            find.Replacement.Text = str(replace_text)
            find.Forward = True
            find.Wrap = 1  # wdFindContinue
            find.MatchCase = False
            find.Execute(Replace=2)  # wdReplaceAll

        # SaveAs2 FileFormat=17 (wdFormatPDF)
        doc.SaveAs2(str(Path(output_path).resolve()), FileFormat=17)
    except Exception as e:
        raise RuntimeError(f"Word 模板处理失败：{e}")
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=0)
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
