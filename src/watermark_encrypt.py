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


# ---------------- 水印排版：不截断 + 可配置 ----------------

# 可用水印字体（Windows 系统字体，加载失败自动退回内置字体）
FONT_FILES = {
    "黑体": "simhei.ttf",
    "楷体": "simkai.ttf",
    "仿宋": "simfang.ttf",
    "宋体": "simsun.ttc",
    "微软雅黑": "msyh.ttc",
}
_font_cache = {}

WM_DEFAULTS = {
    "layout": "auto",     # auto=自动（可读则单行缩号，否则折行）| single=强制单行 | wrap=强制折行
    "fs_max": 0,          # 最大字号，0=按页面自动
    "font": "",           # ""=内置字体；否则 FONT_FILES 的键名
    "color": "#737373",   # 十六进制
    "opacity": 0.25,
    "angle": -1,          # -1=自动（沿页面对角线）；否则角度（度）
    "repeat": 1,          # 1=单个 | 3=三份（防裁剪）| "tile"=平铺 2列×3行
}


def _load_wm_font(name):
    if not name or name not in FONT_FILES:
        return fitz.Font("china-s")
    if name not in _font_cache:
        try:
            _font_cache[name] = fitz.Font(fontfile="C:/Windows/Fonts/" + FONT_FILES[name])
        except Exception:
            _font_cache[name] = fitz.Font("china-s")
    return _font_cache[name]


def _hex_to_rgb(h):
    h = (h or "#737373").lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


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


def plan_layout(text, font, usable_w, usable_h, fs_max, layout="auto"):
    """按布局模式返回 (行列表, 字号)，保证每行都在 usable_w 内。"""
    if layout == "single":  # 强制单行：字号一路缩到放得下
        fs = fs_max
        while fs >= 8 and font.text_length(text, fontsize=fs) > usable_w:
            fs -= 0.5
        return [text], max(fs, 8)
    if layout == "wrap":    # 强制折行：保持大字号，最多 4 行
        fs = fs_max
        while fs >= 10:
            lines = _wrap_by_width(text, font, fs, usable_w)
            if len(lines) <= 4 and len(lines) * fs * 1.4 <= usable_h:
                return lines, fs
            fs -= 1
        return _wrap_by_width(text, font, 10, usable_w), 10
    # auto：单行缩号仍可读就用单行；缩得太小就折行
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
    """在 side×side 方形临时页写水印；返回该页所属文档。"""
    tmp = fitz.open()
    tp = tmp.new_page(width=side, height=side)
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
        for i, ln in enumerate(lines):
            w = font.text_length(ln, fontsize=fs)
            tw.append(fitz.Point(cx - w / 2,
                                 cy - (i - (n - 1) / 2) * line_h + fs * 0.36),
                      ln, font=font, fontsize=fs)
        tw.write_text(tp, color=color, opacity=op)

    rep = cfg.get("repeat", 1)
    if rep == "tile":  # 平铺由 add_watermark 按页面坐标逐格放置，这里只画单块
        rep = 1
    write_block(side / 2, side / 2, side * 0.82, side * 0.70, fs_main, opacity)
    if rep == 3:  # 两侧副本缩小字号，保证完整落在页内
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

    # 方形临时页内容绕中心旋转任意角度后外接框不变，show_pdf_page 不会缩放；
    # 直接用 TextWriter + morph 在本版本 PyMuPDF 会静默丢字，故走此路径

    def place(stamp, cx, cy, side):
        page.show_pdf_page(fitz.Rect(cx - side / 2, cy - side / 2,
                                     cx + side / 2, cy + side / 2),
                           stamp, 0, rotate=angle, overlay=True)

    if cfg.get("repeat") == "tile":
        # 平铺：按页面坐标 2列×3行 逐格放置，每格独立小方形章，保证整格在页内
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


def process_one(src: Path, out: Path, name, open_pw, owner_pw, no_encrypt,
                repeat=1, opacity=0.25, wm_cfg=None):
    cfg = dict(WM_DEFAULTS)
    if repeat in (1, 3, "tile"):
        cfg["repeat"] = repeat
    if opacity:
        cfg["opacity"] = opacity
    if wm_cfg:
        cfg.update(wm_cfg)
    doc = fitz.open(src)
    for page in doc:
        add_watermark(page, name, cfg)

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
    ap.add_argument("--repeat", default="1", choices=["1", "3", "tile"],
                    help="水印份数：1=单个 | 3=三份（防裁剪）| tile=平铺2列×3行")
    ap.add_argument("--opacity", type=float, default=0.25)
    ap.add_argument("--layout", default="auto", choices=["auto", "single", "wrap"],
                    help="布局：auto=自动 | single=强制单行 | wrap=强制折行")
    ap.add_argument("--font", default="", choices=[""] + list(FONT_FILES),
                    help="字体（默认内置）")
    ap.add_argument("--color", default="#737373", help="颜色（十六进制，如 #B22222）")
    ap.add_argument("--angle", type=float, default=-1,
                    help="角度（度）；-1=自动沿对角线")
    ap.add_argument("--fs-max", type=int, default=0, help="最大字号（0=自动）")
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
                           args.no_encrypt,
                           repeat=int(args.repeat) if args.repeat != "tile" else "tile",
                           opacity=args.opacity,
                           wm_cfg={"layout": args.layout, "font": args.font,
                                   "color": args.color, "angle": args.angle,
                                   "fs_max": args.fs_max})
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
