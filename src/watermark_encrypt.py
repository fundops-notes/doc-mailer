#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 打水印 + 双口令加密（一条命令完成）

解决两个问题：
1. 水印主体名过长被截断 —— 自动"换行或缩小字号"，保证完整显示
2. 只有打开口令、没有权限口令 —— 同时设置
     user_pw  = 打开口令（发给 LP）
     owner_pw = 权限口令（禁止修改/批注/复制，不发 LP；默认随机生成）
   加密算法 AES-256，权限只保留：打印、无障碍阅读

用法：
  python watermark_encrypt.py 原件.pdf --name "宁波某某股权投资合伙企业（有限合伙）" --open-pw Abc12345
  python watermark_encrypt.py docs/*.pdf --name "..." --open-pw ... --owner-pw "内部口令!" --outdir out
  python watermark_encrypt.py 原件.pdf --name "..." --open-pw ... --no-encrypt   # 只看水印效果

注意：请用「未打水印的原件」作为输入，不要在已加密版本上二次处理。
"""

import argparse
import glob
import math
import secrets
import sys
from pathlib import Path

import fitz  # PyMuPDF

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ---------------- 水印排版：不截断 ----------------

def _wrap_by_width(text, font, fontsize, max_w):
    """按显示宽度逐字换行（中文公司名没有空格，只能在字符间断行）。"""
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


def plan_layout(text, font, usable_w, usable_h, fs_max):
    """优先单行缩字号；缩得太小( readability 阈值)就保持大字号换行，最多 3 行。"""
    # 方案 A：单行，字号从大到小找第一个能放下的
    fs = fs_max
    while fs >= 10:
        if font.text_length(text, fontsize=fs) <= usable_w:
            break
        fs -= 0.5
    if fs >= max(14, fs_max * 0.55):          # 单行且字号仍可读 → 用它
        return [text], fs
    # 方案 B：保持大字号换行；整块高度放不下再缩
    fs = fs_max
    while fs >= 10:
        lines = _wrap_by_width(text, font, fs, usable_w)
        if len(lines) <= 3 and len(lines) * fs * 1.4 <= usable_h:
            return lines, fs
        fs -= 1
    return _wrap_by_width(text, font, 10, usable_w), 10


def _make_stamp(text, font, side, repeat, opacity):
    """在 side×side 的方形临时页上写好（可能多行的）水印文字块。
    repeat=3 时主水印居中，两侧各一份 55% 字号、更淡的副本（防裁剪）。"""
    tmp = fitz.open()
    tp = tmp.new_page(width=side, height=side)
    usable_w, usable_h = side * 0.82, side * 0.70
    fs_max = max(16, min(40, side * 0.033))
    lines, fs = plan_layout(text, font, usable_w, usable_h, fs_max)

    def write_block(center_x, font_size, op):
        tw = fitz.TextWriter(tp.rect)
        line_h = font_size * 1.35
        n = len(lines)
        for i, ln in enumerate(lines):
            w = font.text_length(ln, fontsize=font_size)
            x = center_x - w / 2
            y = side / 2 - (i - (n - 1) / 2) * line_h + font_size * 0.36
            tw.append(fitz.Point(x, y), ln, font=font, fontsize=font_size)
        tw.write_text(tp, color=(0.45, 0.45, 0.45), opacity=op)

    write_block(side / 2, fs, opacity)
    if repeat > 1:
        for off in (-0.20, 0.20):  # 侧份缩小字号，保证完整落在页内不被裁
            write_block(side / 2 + off * usable_w, fs * 0.55, opacity * 0.6)
    return tmp


def add_watermark(page, text, font, repeat=1, opacity=0.16):
    r = page.rect
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    side = math.ceil(math.hypot(r.width, r.height) * 1.05)
    angle = math.degrees(math.atan2(r.height, r.width))

    # 方形临时页内容绕中心旋转任意角度后外接框不变，show_pdf_page 不会缩放；
    # 直接用 TextWriter + morph 在本版本 PyMuPDF 会静默丢字，故走此路径
    stamp = _make_stamp(text, font, side, repeat, opacity)
    page.show_pdf_page(fitz.Rect(cx - side / 2, cy - side / 2,
                                 cx + side / 2, cy + side / 2),
                       stamp, 0, rotate=angle, overlay=True)


# ---------------- 主流程 ----------------

def _perm_summary(doc):
    """把权限位掩码解成可读 dict（以 LP 用打开口令登录后的视角）。"""
    p = doc.permissions & 0xFFFFFFFF
    return {
        "打印": bool(p & fitz.PDF_PERM_PRINT),
        "修改": bool(p & fitz.PDF_PERM_MODIFY),
        "复制文字": bool(p & fitz.PDF_PERM_COPY),
        "批注": bool(p & fitz.PDF_PERM_ANNOTATE),
        "填表单": bool(p & fitz.PDF_PERM_FORM),
        "页面拼装": bool(p & fitz.PDF_PERM_ASSEMBLE),
    }


def process_one(src: Path, out: Path, name, open_pw, owner_pw, no_encrypt, repeat,
                opacity=0.25):
    font = fitz.Font("china-s")  # 内置简体中文字体，随文件子集嵌入
    doc = fitz.open(src)
    for page in doc:
        add_watermark(page, name, font, repeat=repeat, opacity=opacity)

    if no_encrypt:
        doc.save(out, garbage=3, deflate=True)
        return {"encrypted": False}

    perms = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
    if hasattr(fitz, "PDF_PERM_ACCESSIBILITY"):
        perms |= fitz.PDF_PERM_ACCESSIBILITY
    # 未列入的权限（修改/批注/复制文本/表单填写/页面拼装）全部禁止

    doc.save(out, encryption=fitz.PDF_ENCRYPT_AES_256,
             user_pw=open_pw, owner_pw=owner_pw,
             permissions=perms, garbage=3, deflate=True)
    doc.close()

    # 自检：打开必须需要口令；打开口令可读、权限口令有完整权限
    d = fitz.open(out)
    assert d.needs_pass, "输出文件未加密！"
    u = d.authenticate(open_pw)
    assert u >= 2, "打开口令验证失败"
    wm_text_found = all(pg.get_text().find(name[:6]) >= 0 for pg in d)
    perms = _perm_summary(d)
    d.close()
    d2 = fitz.open(out)
    o = d2.authenticate(owner_pw)
    assert o >= 4, "权限口令验证失败"
    d2.close()
    return {"encrypted": True, "user_ok": bool(u >= 2), "owner_ok": bool(o >= 4),
            "user_perms": perms, "wm_ok": wm_text_found}


def main():
    ap = argparse.ArgumentParser(description="PDF 打水印 + 双口令加密")
    ap.add_argument("inputs", nargs="+", help="PDF 原件（可用通配符）")
    ap.add_argument("--name", required=True, help="水印主体名称（LP 名称）")
    ap.add_argument("--open-pw", required=True, help="打开口令（发给 LP 的）")
    ap.add_argument("--owner-pw", default=None,
                    help="权限口令；不给则随机生成（仅打印在屏幕上，勿发 LP）")
    ap.add_argument("--outdir", default="out_wm", help="输出目录（默认 out_wm）")
    ap.add_argument("--repeat", type=int, default=1, choices=[1, 3],
                    help="水印份数：1=对角线主水印，3=主水印+两侧淡份")
    ap.add_argument("--opacity", type=float, default=0.25)
    ap.add_argument("--no-encrypt", action="store_true", help="只打水印不加密（预览效果用）")
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
                           args.no_encrypt, args.repeat)
        if info.get("encrypted"):
            print(f"[OK] {src.name} -> {out}  水印{'已写入' if info['wm_ok'] else '!!未检出!!'}，"
                  f"打开口令验证{'通过' if info['user_ok'] else '失败'}，"
                  f"权限口令验证{'通过' if info['owner_ok'] else '失败'}")
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
