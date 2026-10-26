#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档分发助手 v3.1 —— 数据层 / 流水线 / 执行引擎（纯逻辑，不含 GUI）

本模块只依赖标准库 + watermark_encrypt，可在无 GUI 环境直接跑测试。
三块内容：
  1. 数据层：lp_data/ 目录化存储、事件/场景/收件人/附件/**产出条目**数据模型、v2 迁移
  2. 流水线：**两种产出条目**（generate 生成文档 / attach 附加文件），
     条目上的处理勾选（转 PDF / 加水印 / 加密，顺序固定）+ 产出粒度
     （per_lp / per_group / all）+ **字段级来源**（每个占位符单独选来源）
  3. 执行引擎：预检 → 条目执行 → 附件组装 → Outlook 草稿 → Run 记录

铁律（不可违反，见 HANDOFF §4–§6）：
  · 权限口令绝不进入邮件正文/标题/任何草稿字段
  · 打开口令 ≠ 权限口令
  · 只建草稿，绝不发送（全库唯一 Outlook 写操作是 mail.Save()）
  · AES-256 加密，权限位仅打印
  · 水印不截断
  · 保存前 subset_fonts()
  · 损坏数据绝不覆盖
"""

import json
import os
import re
import secrets
import shutil
import sys
import threading
from datetime import date, datetime
from pathlib import Path

import fitz
import watermark_encrypt as wm

# ============================================================================
# 基础工具
# ============================================================================

APP_VERSION = "3.9.0"   # 必须与 lp_tool.py 的 APP_VERSION 一致（lp_tool 启动时会断言）

DATA_DIR_NAME = "lp_data"
LEGACY_DATA_FILE = "lp_data.json"
SETTINGS_FILE = "settings.json"
EVENTS_DIR = "events"
SCENARIOS_DIR = "scenarios"
RUNS_DIR = "runs"
RECIPIENTS_DIR = "recipients"   # 收件人库（P4 跨事件复用）


def app_dir() -> Path:
    """程序所在目录（单 exe → exe 目录；源码运行 → 源码目录）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

PW_CHARS = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789"

# 附件路径分隔符：只按 ; ； 和换行（英文逗号是文件名的一部分，如 "X, L.P."）
_PATH_SEPS = ["；", ";", "\n", "\r"]


def default_base_dir() -> Path:
    """程序所在目录（打包成 exe 后是 exe 所在目录）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def gen_pw(n=12):
    """密码学安全随机（random 的梅森旋转可被预测）。"""
    return "".join(secrets.choice(PW_CHARS) for _ in range(n))


def sanitize_filename(s, maxlen=40):
    """非法字符 < > : " / \\ | ? * 用 chr 构造避免字面量转义问题。"""
    for c in map(chr, (60, 62, 58, 34, 47, 92, 124, 63, 42)):
        s = s.replace(c, "_")
    s = "".join(ch for ch in s if ord(ch) >= 32)  # 去控制字符
    s = s.strip().rstrip(chr(46)).strip()          # 结尾的点/空格
    return s[:maxlen] if s else "unnamed"


def pw_fingerprint(pw):
    """口令指纹：只留 8 位哈希，写进执行记录，绝不落明文。"""
    import hashlib
    return hashlib.sha256((pw or "").encode("utf-8")).hexdigest()[:8]


def split_paths(s):
    """拆分附件路径：仅按 ; ；和换行（逗号常出现在中文/英文文件名里，不是分隔符）。"""
    for ch in _PATH_SEPS:
        s = s.replace(ch, ";")
    return [x.strip().strip('"') for x in s.split(";") if x.strip()]


def split_mails(s):
    """拆分邮箱列表。"""
    for ch in ["；", ";", "，", ",", "\n", "\r", "\t"]:
        s = s.replace(ch, ";")
    if "<" not in s:  # 无显示名格式时才把空格当分隔符
        s = s.replace(" ", ";")
    return [x for x in s.split(";") if x]


def read_spreadsheet(path):
    """读 xlsx/csv 为二维列表（第一行是表头）。"""
    p = Path(path)
    if p.suffix.lower() == ".csv":
        import csv
        try:
            with open(p, newline="", encoding="utf-8-sig") as f:
                rows = [row for row in csv.reader(f)]
        except UnicodeDecodeError:  # 中文 Excel 默认导出 GBK
            with open(p, newline="", encoding="gbk", errors="replace") as f:
                rows = [row for row in csv.reader(f)]
        if rows and len(rows[0]) == 1 and ";" in rows[0][0]:  # 分号分隔
            rows = [[c.strip() for c in r[0].split(";")] for r in rows if r]
        return rows
    import openpyxl
    from openpyxl.utils.datetime import from_excel
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows():
        out = []
        for cell in row:
            out.append(_excel_display_text(cell, from_excel))
        rows.append(out)
        if len(rows) > 5000:
            break
    wb.close()
    return rows


def _excel_fmt_has_date(fmt):
    f = str(fmt or "").lower()
    if not f:
        return False
    if "年" in f or "月" in f or "日" in f:
        return True
    # 去掉引号/转义里的普通字符后看是否还有 y/m/d 日期占位
    f2 = re.sub(r'"[^"]*"', "", f)
    f2 = re.sub(r"\\[^a-z]", "", f2)
    return bool(re.search(r"[ymd]", f2)) and not re.search(r"[#0]", f2)


def _excel_display_text(cell, from_excel):
    """单元格 → 文本，尽量按 Excel 里看到的格式（日期不变成数字/带 00:00:00）。"""
    v = getattr(cell, "value", None)
    if v is None:
        return ""
    fmt = str(getattr(cell, "number_format", "") or "")
    # 1) 日期/时间（openpyxl 一般已转成 datetime；个别自定义格式会留数字）
    dt = None
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        dt = v
    elif isinstance(v, (int, float)) and not isinstance(v, bool) and _excel_fmt_has_date(fmt):
        try:
            dt = from_excel(v)
        except Exception:
            dt = None
    if dt is not None:
        has_time = bool(re.search(r"[hHsS]", fmt)) or (
            getattr(dt, "hour", 0) or getattr(dt, "minute", 0) or getattr(dt, "second", 0))
        if "年" in fmt or "月" in fmt or "日" in fmt:
            out = "%d年%d月%d日" % (dt.year, dt.month, dt.day)
        else:
            sep = "-" if "-" in fmt else ("/" if "/" in fmt else "-")
            out = "%04d%s%02d%s%02d" % (dt.year, sep, dt.month, sep, dt.day)
        if has_time and (getattr(dt, "hour", 0) or getattr(dt, "minute", 0)):
            out += " %02d:%02d" % (dt.hour, dt.minute)
            if re.search(r"[sS]", fmt):
                out += ":%02d" % dt.second
        return out
    # 2) 数字：百分比/小数位/千分位
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if "%" in fmt:
            return ("%g" % (v * 100)) + "%"
        dec = 0
        m = re.search(r"0\.([0#]+)", fmt)
        if m:
            dec = len(m.group(1))
        use_comma = "," in fmt
        if dec:
            return format(v, ",.%df" % dec if use_comma else ".%df" % dec)
        if float(v).is_integer():
            iv = int(v)
            return format(iv, ",") if use_comma else str(iv)
        return str(v)
    if isinstance(v, bool):
        return str(v)
    return str(v).strip()


def _cell_text(v):
    """单元格取值转文本：整数的浮点不显示小数尾巴（200000.0 → 200000）。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        if getattr(v, "hour", 0) or getattr(v, "minute", 0):
            return v.strftime("%Y-%m-%d %H:%M")
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def sheet_to_dicts(rows):
    """二维列表 → [{列名: 值}]（跳过全空行）。"""
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for row in rows[1:]:
        rec = {}
        empty = True
        for i, h in enumerate(headers):
            if not h:
                continue
            v = _cell_text(row[i]) if i < len(row) else ""
            rec[h] = v
            if v:
                empty = False
        if not empty:
            out.append(rec)
    return out


HEADER_ALIASES = {
    "name": {"lp名称", "收件人名称", "名称", "主体名称", "客户名称", "lp name", "name"},
    "wm": {"水印名称", "水印"},
    "to": {"收件人", "收件邮箱", "to", "邮箱", "email"},
    "cc": {"抄送", "cc"},
    "bcc": {"密送", "bcc"},
    # v3.4.4：收件人页已不再露出「附件路径」列；导入时该列按「未识别列」
    # 收进自定义字段（fields["附件路径"]），不再写 rec["attachments"]。
    # 仅保留显式的英文/别名列，兼容旧数据导出文件。
    "files": {"额外附件", "attachments", "files"},
    "subject": {"邮件标题", "标题", "主题", "subject"},
    "open_pw": {"打开口令", "打开密码", "open pw"},   # 不收"密码/口令"这种歧义列名
    "owner_pw": {"权限口令", "权限密码", "owner pw"},
    "merge_group": {"合并组", "merge group", "merge_group", "group", "组"},
    "from_addr": {"发件人", "代发地址", "from", "sent on behalf of"},
}


def parse_import_rows(rows):
    """rows: 二维列表，第一行为表头。返回 (收件人记录列表, 未识别的列名列表)。

    v2 兼容实现：口令留空时保留空串（导入合并时不清掉已有口令）。
    """
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
        if not cell("name") or cell("name").startswith("示例："):
            continue  # 跳过空行与导入模板里的示例占位行
        lps.append({
            "name": cell("name"),
            "wm": cell("wm") or cell("name"),
            "to": cell("to"),
            "cc": cell("cc"),
            "bcc": cell("bcc"),
            "subject": cell("subject"),
            "files": cell("files"),
            "open_pw": cell("open_pw"),
            "owner_pw": cell("owner_pw"),
            "merge_group": cell("merge_group"),
            "from_addr": cell("from_addr"),
        })
    return lps, unknown


DATE_HEADER_RE = re.compile(r"date|due|日期|时间|截止|到期|期限", re.I)


def coerce_date_like_value(header, value):
    """日期类列：把 Excel 的日期序列号（如 46193）转成 YYYY-MM-DD。

    有的模板列格式是 General，openpyxl 只能拿到数字，Excel 里看着也像数字，
    但列名是 Due_Date / 日期 时用户意图显然是日期，导入时统一转成可读日期。
    """
    if value is None or value == "":
        return value
    if not DATE_HEADER_RE.search(str(header or "")):
        return value
    if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
        return value.strftime("%Y-%m-%d")
    txt = str(value).strip()
    if not re.fullmatch(r"\d{5}(\.\d+)?", txt):
        return value
    try:
        num = float(txt)
        if 20000 <= num <= 90000:
            from openpyxl.utils.datetime import from_excel
            return from_excel(num).strftime("%Y-%m-%d")
    except Exception:
        pass
    return value


def parse_import_rows_v3(rows):
    """v3 收件人导入。

    识别不了的列 → 收件人的自定义字段（供文档映射用）。
    返回 (recipients, unknown_cols, field_cols)。
    """
    if not rows:
        return [], [], []
    header = ["" if h is None else str(h).strip() for h in rows[0]]
    lower = [h.lower() for h in header]
    col_map = {}
    unknown = []
    field_cols = []   # [(列下标, 原始列名)]
    for idx, h in enumerate(lower):
        hit = None
        for key, aliases in HEADER_ALIASES.items():
            if h in aliases:
                hit = key
                break
        if hit:
            col_map[hit] = idx
        elif h:
            unknown.append(header[idx])
            field_cols.append((idx, header[idx]))
    recips = []
    for row in rows[1:]:
        def cell(key):
            i = col_map.get(key)
            if i is None or i >= len(row):
                return ""
            v = row[i]
            return "" if v is None else str(v).strip()
        name = cell("name")
        if not name or name.startswith("示例："):
            continue
        fields = {}
        for idx, col in field_cols:
            if idx < len(row):
                v = row[idx]
                if v is not None and str(v).strip():
                    fields[col] = str(v).strip()
        recips.append(normalize_recipient({
            "name": name,
            "wm": cell("wm") or name,
            "to": cell("to"),
            "cc": cell("cc"),
            "bcc": cell("bcc"),
            "subject": cell("subject"),
            "files": cell("files"),
            "open_pw": cell("open_pw"),
            "owner_pw": cell("owner_pw"),
            "merge_group": cell("merge_group"),
            "from_addr": cell("from_addr"),
            "fields": fields,
        }))
    return recips, unknown, [c for _i, c in field_cols]


def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _stamp():
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def _new_id(prefix):
    return "%s-%s-%s" % (prefix, _stamp(), secrets.token_hex(3))


# ============================================================================
# 数据模型
# ============================================================================

DEFAULT_FILENAME_TPL = "{{原件名}}_{{LP名称}}"

# 演示数据（均为虚构，邮箱用保留域名，可直接跑通全流程）
MOCK_LPS = [
    ["苏州启衡科技有限公司", "", "ir@qiheng-demo.cn", "fin@qiheng-demo.cn", "", "", "", ""],
    ["宁波晨曦汐诺供应链股份有限公司", "", "contact@chenxi-demo.cn", "", "", "", "Demo-Open-88", ""],
    ["新疆乌鲁木齐市泓远红杉智能科技有限公司", "", "admin@hongyuan-demo.cn", "ir@hongyuan-demo.cn", "audit@hongyuan-demo.cn", "", "", ""],
    ["上海临港科创一期产业投资有限公司", "", "lg@chuangxin-demo.cn", "", "", "临港科创2026年度交付文件", "", ""],
    ["珠海横琴粤澳深谋咨询服务有限公司", "", "gp@henqin-demo.cn", "fin@henqin-demo.cn", "", "", "", "Demo-Owner-66"],
    ["北京中关军融创启信息技术有限公司", "演示专用水印", "zg@rongchuang-demo.cn", "", "", "", "", ""],
    ["深圳前海深港现代服务业合作区汇金聚富实业有限公司", "", "huijin@qianhai-demo.cn", "", "", "", "", ""],
    ["杭州西湖区云栖数字经济研究院", "", "yunqi@xihu-demo.cn", "ir@xihu-demo.cn", "", "", "", ""],
]
MOCK_HEADERS = ["收件人名称", "水印名称", "收件人", "抄送", "密送", "邮件标题",
                "打开口令", "权限口令"]

# v3 默认邮件模板（与 v2 保持一致的措辞）
DEFAULT_SUBJECT_V3 = "机密文件（{{收件人名称}} 专属）"
DEFAULT_BODY_V3 = """<div style="font-family:微软雅黑,Segoe UI,Arial;font-size:14px;color:#333;line-height:1.8;">
<p><b>{{收件人名称}}</b> 您好：</p>
<p>附件为贵司专属的加密文件（已加盖专属水印），请使用以下打开口令查阅：</p>
<p style="padding:10px 16px;background:#f5f6f8;border-left:3px solid #4a6fa5;border-radius:4px;">
<b>打开口令：{{打开口令}}</b></p>
<p style="color:#888;font-size:12px;">本文件带有贵司专属水印标识，仅供贵司内部使用，请妥善保管、请勿转发。</p>
<p>如有任何疑问，请随时与我们联系。<br><br>此致<br>【公司/团队名称】<br>{{日期}}</p>
</div>"""

DEFAULT_RECIPIENT = {
    "name": "", "to": "", "cc": "", "bcc": "",
    "wm": "", "group": "", "subject": "",
    "fields": {},
    "open_pw": "", "owner_pw": "",
    "attachments": [],
}

DEFAULT_ATTACHMENT = {
    "name": "", "path": "", "scope": "all", "targets": [], "group": "",
}

DEFAULT_MAIL = {
    "subject": "", "body": "", "from_addr": "", "cc_default": "", "bcc_default": "",
}

DEFAULT_OPTIONS = {
    "merge_groups": True,
    "send_mode": "per_group",   # 发送模式：per_lp 每人一份 / per_group 按组合并 / all 全部合并
    "gp_owner_mode": False,
    "gp_owner_pw": "",
    "outdir": "",
    "filename_tpl": DEFAULT_FILENAME_TPL,
    "wm_settings": {},
    "constants": {},
}

# ---------- 发送模式 ----------
# 事件级开关：决定「几封邮件 / 几份文档」。
#   per_lp    每人一份：每位收件人一封邮件、各自一份文档（最常见）
#   per_group 按组合并：同一合并组的收件人共用一封邮件，附件包合并
#   all       全部合并：所有收件人共用一封邮件、一份文档
SEND_MODES = [
    ("per_lp", "每人一份", "每位收件人各自一封邮件、各自一份文档"),
    ("per_group", "按组合并", "同一合并组的收件人共用一封邮件，附件包合并"),
    ("all", "全部合并", "所有收件人共用一封邮件、一份文档"),
]
SEND_MODE_KEYS = [k for k, _l, _d in SEND_MODES]
SEND_MODE_LABELS = {k: l for k, l, _d in SEND_MODES}
SEND_MODE_DESC = {k: d for k, _l, d in SEND_MODES}
SEND_MODE_DEFAULT = "per_group"


def send_mode_of(event):
    """取事件的发送模式。

    认 options.send_mode；没有就退回旧的 options.merge_groups 布尔（v2 迁移数据）。
    """
    opts = ((event or {}).get("options") or {})
    m = str(opts.get("send_mode") or "").strip()
    if m in SEND_MODE_KEYS:
        return m
    if "merge_groups" in opts:
        return "per_group" if opts.get("merge_groups") else "per_lp"
    return SEND_MODE_DEFAULT


DEFAULT_SETTINGS = {
    "version": 3,
    "wm_settings": {},
    "gp_owner_mode": False,
    "gp_owner_pw": "",
    "outdir": "",
    "filename_tpl": DEFAULT_FILENAME_TPL,
    "auto_update": True,
    "field_templates": [
        {"name": "催缴场景字段", "columns": ["Call_Amount", "Unfunded", "Due_Date"]},
        {"name": "对账场景字段", "columns": ["Balance", "Paid_Amount", "Statement_Date"]},
    ],
    "mail_subject": "",
    "mail_body": "",
    "from_addr": "",
}

# ============================================================================
# 产出条目（v3.1 重构）
# ============================================================================
# 流水线上每一行 = 一个要附上的东西，只有两种：
#   generate 生成文档：Word 模板 → .docx →（可选）PDF →（可选）水印 →（可选）加密
#   attach   附加文件：磁盘上已有的文件 →（可选，仅 PDF）水印 →（可选，仅 PDF）加密
# 「转 PDF / 加水印 / 加密」是条目上的勾选项，不是独立步骤（顺序固定，用户不用排）。
ITEM_KINDS = ("generate", "attach")

ITEM_KIND_LABELS = {
    "generate": "Word-MailMerge",
    "attach": "附加文件",
}

ITEM_KIND_DESC = {
    "generate": "Word 邮件合并：按占位符把收件人字段替换进模板，可转 PDF / 加水印 / 加密",
    "attach": "把磁盘上已有的文件放进附件包",
}

# 产出粒度：这一条产出几份文件
GRANULARITIES = ("per_lp", "per_group", "all")

GRANULARITY_LABELS = {
    "per_lp": "每人一份",
    "per_group": "每组合并一份",
    "all": "全员共用一份",
}

GRANULARITY_DESC = {
    "per_lp": "每位收件人各得一份，用他自己的字段、水印和口令",
    "per_group": "同一合并组共用一份，用组内第一个人的字段",
    "all": "所有收件人共用同一份（口令也共享）",
}

# PDF 权限位（v3.1 起可配置；默认仍为「仅打印」，保持既有铁律的默认行为）
PERMISSIONS = ("print", "edit", "copy")

PERMISSION_LABELS = {
    "print": "允许打印",
    "edit": "允许编辑",
    "copy": "允许复制内容",
}

PERMISSION_DESC = {
    "print": "收件人可以打印这份文件",
    "edit": "收件人打开后可以修改内容",
    "copy": "收件人可以复制里面的文字和图片",
}

DEFAULT_PERMISSIONS = ["print"]

# ---------------------------------------------------------------------------
# 附加文件条目的「路径来源」
# ---------------------------------------------------------------------------
# 以前附加文件条目只能手填一串固定路径 → 每个人拿到的都一样。
# 现在可以按收件人取路径，让「每人一份不同的文件」在流水线上就能配置：
#   manual              手填路径（默认，行为与旧版一致）
#   recipient.attachments 收件人清单的「附件路径」列（导入/手填的那一列）
#   fields.<列名>        收件人自定义字段里的某一列
#   sheet.<列名>         外部数据表里的某一列
PATH_SOURCE_MANUAL = "manual"
PATH_SOURCE_RECIPIENT = "recipient.attachments"
PATH_SOURCE_FIELD_PREFIX = "fields."
PATH_SOURCE_SHEET_PREFIX = "sheet."

DEFAULT_ITEM = {
    "id": "", "kind": "generate", "name": "", "enabled": True,
    "params": {}, "granularity": "per_group", "mapping": {}, "sheet": {},
}

# 占位符来源种类（**字段级**，不再是步骤级）
SOURCE_RECIPIENT = "recipient"   # 收件人字段
SOURCE_CONSTANT = "constant"     # 事件常量
SOURCE_SHEET = "sheet"           # 外部 Excel


def normalize_recipient(rec):
    """把任意来源的收件人记录规范成 v3 结构（容错：坏数据不抛异常）。"""
    out = dict(DEFAULT_RECIPIENT)
    out["fields"] = {}
    out["attachments"] = []
    if not isinstance(rec, dict):
        return out
    for k in ("name", "to", "cc", "bcc", "open_pw", "owner_pw", "subject"):
        v = rec.get(k)
        if isinstance(v, str):
            out[k] = v
    # 水印文字：v2 用 wm，v3 概念名 wm_text，两者都收，统一存 wm
    for k in ("wm", "wm_text"):
        v = rec.get(k)
        if isinstance(v, str) and v:
            out["wm"] = v
            break
    # 合并组：v2 用 merge_group，v3 用 group
    for k in ("group", "merge_group"):
        v = rec.get(k)
        if isinstance(v, str) and v:
            out["group"] = v
            break
    fields = rec.get("fields")
    if isinstance(fields, dict):
        for k, v in fields.items():
            if v is None:
                continue
            out["fields"][str(k)] = str(v)
    atts = rec.get("attachments")
    if isinstance(atts, list):
        out["attachments"] = [str(a).strip() for a in atts
                              if isinstance(a, str) and a.strip()]
    # v2 的 files 字段（分号分隔的附件路径）→ attachments
    files = rec.get("files")
    if isinstance(files, str) and files.strip():
        for p in split_paths(files):
            if p not in out["attachments"]:
                out["attachments"].append(p)
    return out


def normalize_attachment(att):
    out = dict(DEFAULT_ATTACHMENT)
    out["targets"] = []
    if not isinstance(att, dict):
        return out
    for k in ("name", "path", "group"):
        v = att.get(k)
        if isinstance(v, str):
            out[k] = v
    scope = att.get("scope")
    if scope in ("all", "specific", "group"):
        out["scope"] = scope
    targets = att.get("targets")
    if isinstance(targets, list):
        out["targets"] = [str(t) for t in targets if isinstance(t, str) and t.strip()]
    elif isinstance(targets, str) and targets.strip():
        out["targets"] = split_paths(targets)
    return out


def normalize_item(it):
    """把任意来源的条目配置规范成 v3.1 结构（容错：坏数据不抛异常）。"""
    out = {
        "id": "", "kind": "generate", "name": "", "enabled": True,
        "params": {}, "granularity": "per_group", "mapping": {}, "sheet": {},
    }
    if not isinstance(it, dict):
        it = {}
    k = it.get("kind")
    out["kind"] = k if k in ITEM_KINDS else "generate"
    for k in ("id", "name"):
        v = it.get(k)
        if isinstance(v, str):
            out[k] = v
    if not out["id"]:
        out["id"] = _new_id("it")
    en = it.get("enabled")
    out["enabled"] = bool(en) if en is not None else True
    g = it.get("granularity")
    # v3.5.0：默认按分组产出（同组一份、未分组各自一份）
    out["granularity"] = g if g in GRANULARITIES else "per_group"

    params = it.get("params")
    params = params if isinstance(params, dict) else {}
    p = {
        "template_path": "",       # kind=generate
        "files": [],               # kind=attach
        "path_source": PATH_SOURCE_MANUAL,   # kind=attach：路径从哪来
        "to_pdf": False,
        "watermark": False,
        "encrypt": False,
        "permissions": list(DEFAULT_PERMISSIONS),
        "filename_tpl": "",       # kind=generate：这一条统一使用的产物文件名模板
    }
    v = params.get("template_path")
    if isinstance(v, str):
        p["template_path"] = v
    ps = params.get("path_source")
    if isinstance(ps, str) and ps.strip():
        p["path_source"] = ps.strip()
    fl = params.get("files")
    if isinstance(fl, list):
        p["files"] = [str(x).strip() for x in fl
                      if isinstance(x, (str, Path)) and str(x).strip()]
    elif isinstance(fl, str) and fl.strip():
        p["files"] = split_paths(fl)
    for b in ("to_pdf", "watermark", "encrypt"):
        if b in params:
            p[b] = bool(params[b])
    if params.get("input_type") in ("pdf", "file", "excel"):
        p["input_type"] = params["input_type"]
    path_map = params.get("path_map")
    if isinstance(path_map, dict):
        p["path_map"] = {str(k): split_paths(v) if isinstance(v, str) else
                         [str(x) for x in v] for k, v in path_map.items()
                         if isinstance(v, (str, list))}
    if params.get("preserve_name"):
        p["preserve_name"] = True
    if isinstance(params.get("target_scope"), dict):
        p["target_scope"] = dict(params["target_scope"])
    ftpl = params.get("filename_tpl")
    if isinstance(ftpl, str):
        p["filename_tpl"] = ftpl.strip()
    perms = params.get("permissions")
    if isinstance(perms, list):
        # 注意：用户主动清空权限位时保留空列表（= 只能打开查看）；
        # 没配 permissions 键时才用默认的「仅打印」。
        p["permissions"] = [x for x in perms if x in PERMISSIONS]
    out["params"] = p

    # 字段映射：{占位符: 来源键}，来源键见 field_source_options()
    mp = it.get("mapping")
    mp = mp if isinstance(mp, dict) else {}
    mapping = {}
    mf = mp.get("fields") if isinstance(mp.get("fields"), dict) else mp
    for kk, vv in (mf or {}).items():
        if isinstance(kk, str) and isinstance(vv, str):
            mapping[kk] = vv
    out["mapping"] = mapping

    # 外部数据表（只有映射里用到 sheet. 来源时才需要）
    sh = it.get("sheet")
    sh = sh if isinstance(sh, dict) else {}
    sheet = {"path": "", "key_col": "", "match_col": "name"}
    for kk in ("path", "key_col", "match_col"):
        if isinstance(sh.get(kk), str):
            sheet[kk] = sh[kk]
    out["sheet"] = sheet
    return out


def new_item(kind="generate", name=""):
    """新建一个条目。"""
    variant = kind if kind in ("pdf", "file", "excel") else None
    it = normalize_item({"kind": "attach" if variant else kind, "name": name,
                         "params": {"input_type": variant} if variant else {}})
    if variant == "excel":
        it["params"]["path_source"] = "item.paths"
    if not it["name"]:
        it["name"] = {"generate": "新生成文档",
                       "attach": "新附加文件", "pdf": "PDF 文件",
                       "file": "非 PDF 文件", "excel": "Excel 按 LP 匹配"}.get(kind, "新条目")
    return it


def item_input_type(item):
    if item.get("kind", "generate") == "generate":
        return "generate"
    params = item.get("params") or {}
    if params.get("input_type") in ("pdf", "file", "excel"):
        return params["input_type"]
    if params.get("path_source", PATH_SOURCE_MANUAL) != PATH_SOURCE_MANUAL:
        return "excel"
    files = params.get("files") or []
    return "pdf" if files and all(Path(p).suffix.lower() == ".pdf" for p in files) else "file"


def item_type_label(item):
    return {"generate": "Word MailMerge", "pdf": "PDF 文件", "file": "非 PDF 文件",
            "excel": "Excel 按 LP 匹配"}[item_input_type(item)]


def upgrade_pipeline_event(event):
    """Expose old static attachments in the new pipeline, without writing on read."""
    import copy
    ev = copy.deepcopy(event)
    if ev.get("pipeline_ui_version") == "3.9":
        return ev
    for index, attachment in enumerate(ev.get("attachments") or []):
        path = attachment.get("path") or ""
        if not path:
            continue
        item = new_item("pdf" if Path(path).suffix.lower() == ".pdf" else "file",
                        "原统一附件 · " + Path(path).name)
        item["id"] = "%s-legacy-common-%d" % (ev.get("id", "event"), index)
        item["params"].update(files=[path], preserve_name=True, target_scope=attachment)
        item["granularity"] = "all" if attachment.get("scope", "all") == "all" else "per_lp"
        ev.setdefault("pipeline", []).append(item)
    ev["attachments"] = []
    if any(r.get("attachments") for r in ev.get("recipients") or []) and not event_uses_recipient_paths(ev):
        item = new_item("excel", "原收件人专属文件")
        item["id"] = "%s-legacy-recipient-files" % ev.get("id", "event")
        item["params"]["path_source"] = PATH_SOURCE_RECIPIENT
        ev.setdefault("pipeline", []).append(item)
    ev["pipeline_ui_version"] = "3.9"
    return ev


def item_uses_sheet(item):
    """这条目的字段映射（或路径来源）里是否用到了外部数据表。"""
    mp = (item or {}).get("mapping") or {}
    if any(str(v).startswith("sheet.") for v in mp.values()):
        return True
    ps = str(((item or {}).get("params") or {}).get("path_source") or "")
    return ps.startswith(PATH_SOURCE_SHEET_PREFIX)


def normalize_mail(mail):
    out = dict(DEFAULT_MAIL)
    if not isinstance(mail, dict):
        return out
    for k in DEFAULT_MAIL:
        v = mail.get(k)
        if isinstance(v, str):
            out[k] = v
    return out


def normalize_options(opts):
    out = dict(DEFAULT_OPTIONS)
    out["wm_settings"] = {}
    out["constants"] = {}
    if not isinstance(opts, dict):
        return out
    for k in ("merge_groups", "gp_owner_mode"):
        if k in opts:
            out[k] = bool(opts[k])
    for k in ("gp_owner_pw", "outdir", "filename_tpl"):
        v = opts.get(k)
        if isinstance(v, str):
            out[k] = v
    if not out["filename_tpl"]:
        out["filename_tpl"] = DEFAULT_FILENAME_TPL
    ws = opts.get("wm_settings")
    if isinstance(ws, dict):
        out["wm_settings"] = dict(ws)
    consts = opts.get("constants")
    if isinstance(consts, dict):
        out["constants"] = {str(k): ("" if v is None else str(v))
                            for k, v in consts.items()}
    sm = str(opts.get("send_mode") or "").strip()
    if sm in SEND_MODE_KEYS:
        out["send_mode"] = sm
        out["merge_groups"] = (sm != "per_lp")
    elif "merge_groups" in opts:
        out["send_mode"] = "per_group" if opts.get("merge_groups") else "per_lp"
    return out


def normalize_event(ev):
    out = {
        "id": "", "name": "", "biz_type": "", "due_date": "", "notes": "",
        "created_at": "", "updated_at": "", "status": "",
        "scenario_id": "", "scenario_name": "",
        "recipients": [], "pipeline": [], "attachments": [],
        "mail": dict(DEFAULT_MAIL), "options": dict(DEFAULT_OPTIONS),
        "process_results": {},
        "mail_overrides": {},
    }
    if not isinstance(ev, dict):
        return out
    for key in ("pipeline_ui_version", "process_signature"):
        if isinstance(ev.get(key), str):
            out[key] = ev[key]
    pr = ev.get("process_results")
    if isinstance(pr, dict):
        out["process_results"] = pr
    po = ev.get("process_overrides")
    if isinstance(po, dict):
        out["process_overrides"] = po
    mo = ev.get("mail_overrides")
    if isinstance(mo, dict):
        out["mail_overrides"] = {
            str(k): {"subject": str(v.get("subject") or ""),
                     "body": str(v.get("body") or "")}
            for k, v in mo.items() if isinstance(v, dict)
        }
    for k in ("id", "name", "biz_type", "due_date", "notes",
              "created_at", "updated_at", "status", "scenario_id", "scenario_name"):
        v = ev.get(k)
        if isinstance(v, str):
            out[k] = v
    if not out["id"]:
        out["id"] = _new_id("evt")
    recips = ev.get("recipients")
    if isinstance(recips, list):
        out["recipients"] = [normalize_recipient(r) for r in recips]
    pipe = ev.get("pipeline")
    if isinstance(pipe, list):
        out["pipeline"] = [normalize_item(s) for s in pipe]
    atts = ev.get("attachments")
    if isinstance(atts, list):
        out["attachments"] = [normalize_attachment(a) for a in atts]
    out["mail"] = normalize_mail(ev.get("mail"))
    out["options"] = normalize_options(ev.get("options"))
    return out


def normalize_scenario(sc):
    out = {
        "id": "", "name": "", "description": "",
        "created_at": "", "updated_at": "",
        "recipients": [], "pipeline": [], "attachments": [],
        "mail": dict(DEFAULT_MAIL), "options": dict(DEFAULT_OPTIONS),
    }
    if not isinstance(sc, dict):
        return out
    for k in ("id", "name", "description", "created_at", "updated_at"):
        v = sc.get(k)
        if isinstance(v, str):
            out[k] = v
    if not out["id"]:
        out["id"] = _new_id("scn")
    recips = sc.get("recipients")
    if isinstance(recips, list):
        out["recipients"] = [normalize_recipient(r) for r in recips]
    pipe = sc.get("pipeline")
    if isinstance(pipe, list):
        out["pipeline"] = [normalize_item(s) for s in pipe]
    atts = sc.get("attachments")
    if isinstance(atts, list):
        out["attachments"] = [normalize_attachment(a) for a in atts]
    out["mail"] = normalize_mail(sc.get("mail"))
    out["options"] = normalize_options(sc.get("options"))
    return out


def new_event(name="未命名事件", scenario=None):
    """新建发送事件。若给了场景，则复制一份场景配置（复制式，改场景不影响已建事件）。"""
    ev = normalize_event({
        "name": name,
        "created_at": _now_str(),
        "updated_at": _now_str(),
        "status": "草稿",
    })
    if scenario:
        ev["scenario_id"] = scenario.get("id", "")
        ev["scenario_name"] = scenario.get("name", "")
        ev["recipients"] = [normalize_recipient(r) for r in scenario.get("recipients", [])]
        ev["pipeline"] = [normalize_item(s) for s in scenario.get("pipeline", [])]
        ev["attachments"] = [normalize_attachment(a) for a in scenario.get("attachments", [])]
        ev["mail"] = normalize_mail(scenario.get("mail"))
        ev["options"] = normalize_options(scenario.get("options"))
    return ev


def new_scenario(name="新建场景", description=""):
    return normalize_scenario({
        "name": name, "description": description,
        "created_at": _now_str(), "updated_at": _now_str(),
    })


def clone_event(ev, new_name=None):
    """复制事件（新 ID，收件人/流水线/附件深拷贝）。"""
    data = json.loads(json.dumps(ev, ensure_ascii=False))
    data["id"] = _new_id("evt")
    data["name"] = new_name or (ev.get("name", "事件") + " 副本")
    data["created_at"] = _now_str()
    data["updated_at"] = _now_str()
    data["status"] = "草稿"
    data["runs"] = []
    for st in data.get("pipeline", []):
        st["id"] = _new_id("st")
    return normalize_event(data)


def sync_scenario_to_event(scenario, event):
    """把场景配置同步进事件（复制式，用户显式触发）。

    只同步「怎么做」：流水线/附件/邮件模板/选项/收件人模板。
    收件人按名称合并：已存在的收件人保留其口令与自定义字段。
    返回 (event, 变更说明列表)。
    """
    ev = normalize_event(json.loads(json.dumps(event, ensure_ascii=False)))
    notes = []
    ev["pipeline"] = [normalize_item(s) for s in scenario.get("pipeline", [])]
    ev["attachments"] = [normalize_attachment(a) for a in scenario.get("attachments", [])]
    ev["mail"] = normalize_mail(scenario.get("mail"))
    ev["options"] = normalize_options(scenario.get("options"))
    notes.append("流水线：%d 步" % len(ev["pipeline"]))
    notes.append("附件：%d 个" % len(ev["attachments"]))
    # 收件人：按名称合并，保留已有口令
    existing = {r.get("name", ""): r for r in ev.get("recipients", [])}
    added = 0
    for rec in scenario.get("recipients", []):
        src = normalize_recipient(rec)
        name = src.get("name", "")
        if not name:
            continue
        cur = existing.get(name)
        if cur is None:
            ev["recipients"].append(src)
            added += 1
        else:
            # 只补空缺，不覆盖已有口令与自定义字段
            for k in ("to", "cc", "bcc", "wm", "group"):
                if not cur.get(k) and src.get(k):
                    cur[k] = src[k]
            for k, v in src.get("fields", {}).items():
                cur.setdefault("fields", {}).setdefault(k, v)
    notes.append("收件人：新增 %d 个，其余按名称合并（已有口令保留）" % added)
    ev["updated_at"] = _now_str()
    return ev, notes


# ============================================================================
# 收件人库（P4：跨事件复用）
# ============================================================================
#
# 收件人库是一份**独立的常用收件人清单**，用来在多个事件之间复用同一批收件人。
# 它不替代事件里的收件人：事件仍是「这一批发给谁」的唯一真相来源；
# 库只是「常用的那批人」，靠显式的「存入 / 添加」按钮流动，避免意外改动已发批次。
#
# 铁律：合并时**绝不覆盖目标已有的打开口令与权限口令**（见 merge_recipient_into_list）。


def normalize_pool_recipient(rec):
    """收件人库记录规范化：收件人结构 + id + updated_at。"""
    out = dict(normalize_recipient(rec))
    out["id"] = ""
    out["updated_at"] = ""
    out["tags"] = []
    if isinstance(rec, dict):
        for k in ("id", "updated_at"):
            v = rec.get(k)
            if isinstance(v, str):
                out[k] = v
        tags = rec.get("tags")
        if isinstance(tags, list):
            seen = set()
            for t in tags:
                t = str(t).strip()
                if t and t not in seen:
                    seen.add(t)
                    out["tags"].append(t)
    if not out["id"]:
        out["id"] = _new_id("rcp")
    return out


def new_pool_recipient(rec=None):
    """新建一条收件人库记录。"""
    out = normalize_pool_recipient(rec or {})
    out["id"] = _new_id("rcp")
    out["updated_at"] = _now_str()
    return out


def merge_recipient_into_list(recipients, rec):
    """把 rec 并入 recipients 列表（按收件人名称匹配）。

    返回 (recipients, added, updated)：
      · 没有同名 → 追加，added+1
      · 有同名   → **只补空缺**：邮箱/抄送/密送/水印/组/标题/附件路径/自定义字段的空值；
                   **绝不覆盖目标已有的打开口令与权限口令**（口径与场景同步一致）。

    为什么要这样：口令是一次性发出的凭据，库里那份可能是旧值；
    若覆盖，等于把已经发出去的批次的口令改掉。
    """
    recipients = recipients if isinstance(recipients, list) else []
    src = normalize_recipient(rec)
    name = src.get("name", "")
    if not name:
        return recipients, 0, 0
    cur = None
    for r in recipients:
        if (r or {}).get("name", "") == name:
            cur = r
            break
    if cur is None:
        recipients.append(src)
        return recipients, 1, 0
    for k in ("to", "cc", "bcc", "wm", "group", "subject"):
        if not cur.get(k) and src.get(k):
            cur[k] = src[k]
    for k, v in src.get("fields", {}).items():
        cur.setdefault("fields", {}).setdefault(k, v)
    for p in src.get("attachments") or []:
        cur.setdefault("attachments", [])
        if p not in cur["attachments"]:
            cur["attachments"].append(p)
    return recipients, 0, 1


def pool_merge_recipient(pool, rec):
    """按名称把一条收件人并入库。

    返回 (pool, 新增的那条 or None, 是否是更新)。
    调用方拿到"新增的那条"才需要落盘 —— 更新已有条目时**不动它原有的口令**。
    """
    pool = pool if isinstance(pool, list) else []
    src = normalize_pool_recipient(rec)
    name = src.get("name", "")
    if not name:
        return pool, None, False
    for r in pool:
        if (r or {}).get("name", "") == name:
            merged, _a, _u = merge_recipient_into_list(pool, src)
            # 已存在：标签取并集（导入/合并不丢标签）
            cur = r
            tags = list(cur.get("tags") or [])
            for t in src.get("tags") or []:
                if t not in tags:
                    tags.append(t)
            cur["tags"] = tags
            return merged, None, True
    pool.append(src)
    return pool, src, False


def merge_recipients_into_event(event, records):
    """把若干收件人并入事件（按名称合并，保留目标已有口令）。

    返回 (event, 新增数, 更新数)。
    """
    ev = normalize_event(event)
    added = updated = 0
    for rec in records or []:
        ev["recipients"], a, u = merge_recipient_into_list(ev["recipients"], rec)
        added += a
        updated += u
    ev["updated_at"] = _now_str()
    return ev, added, updated


# ============================================================================
# 数据层：目录化存储
# ============================================================================

def read_json_file(path):
    """读 JSON。返回 (data, error)；error 非 None 表示文件存在但读不了。"""
    p = Path(path)
    if not p.exists():
        return None, None
    try:
        raw = p.read_text(encoding="utf-8-sig")  # 容忍 BOM
    except Exception as e:
        return None, str(e)
    if not raw.strip():
        return None, "文件为空"
    try:
        return json.loads(raw), None
    except Exception as e:
        return None, str(e)


def write_json_file(path, data):
    """原子写：先写 .tmp 再替换；旧文件轮转 .bak。返回 (ok, error)。"""
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(p.name + ".tmp")
    bak = p.with_name(p.name + ".bak")
    try:
        raw = json.dumps(data, ensure_ascii=False, indent=2)
        tmp.write_text(raw, encoding="utf-8")
        if p.exists():
            p.replace(bak)
        os.replace(tmp, p)
        return True, None
    except Exception as e:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
        return False, str(e)


class Store:
    """lp_data/ 目录化存储。

    结构：
      lp_data/settings.json
      lp_data/events/<event_id>.json
      lp_data/scenarios/<scenario_id>.json
      lp_data/runs/<event_id>-<ts>.json
    """

    def __init__(self, base_dir=None):
        self.base = Path(base_dir) if base_dir else default_base_dir()
        self.data_dir = self.base / DATA_DIR_NAME
        self.settings_file = self.data_dir / SETTINGS_FILE
        self.events_dir = self.data_dir / EVENTS_DIR
        self.scenarios_dir = self.data_dir / SCENARIOS_DIR
        self.runs_dir = self.data_dir / RUNS_DIR
        self.recipients_dir = self.data_dir / RECIPIENTS_DIR
        self.legacy_file = self.base / LEGACY_DATA_FILE

    # ----- 目录 -----

    def ensure_dirs(self):
        for d in (self.data_dir, self.events_dir, self.scenarios_dir,
                  self.runs_dir, self.recipients_dir):
            d.mkdir(parents=True, exist_ok=True)

    def exists(self):
        return self.data_dir.exists()

    # ----- 设置 -----

    def load_settings(self):
        data, err = read_json_file(self.settings_file)
        if err is not None:
            # 损坏：改名留证，绝不覆盖（防口令丢失）
            bak = self._quarantine(self.settings_file, "settings.corrupt")
            self.last_error = "settings.json 读取失败：%s%s" % (
                err, ("（原文件已备份为 %s）" % bak.name) if bak else "")
            return dict(DEFAULT_SETTINGS), self.last_error
        if data is None:
            return dict(DEFAULT_SETTINGS), None
        out = dict(DEFAULT_SETTINGS)
        if isinstance(data, dict):
            for k, v in data.items():
                out[k] = v
        out.setdefault("field_templates", list(DEFAULT_SETTINGS["field_templates"]))
        return out, None

    def save_settings(self, data):
        self.ensure_dirs()
        ok, err = write_json_file(self.settings_file, data)
        return ok, err

    def get_settings(self):
        """读全局设置（出错时回落到默认值并记 last_error）。"""
        data, err = self.load_settings()
        if err:
            self.last_error = err
        return data

    # ----- 事件 -----

    def event_path(self, event_id):
        return self.events_dir / ("%s.json" % sanitize_filename(event_id, 80))

    def list_events(self):
        """返回 [event dict]，按更新时间倒序。损坏文件跳过并记录。"""
        out = []
        if not self.events_dir.exists():
            return out
        for p in sorted(self.events_dir.glob("*.json")):
            data, err = read_json_file(p)
            if err is not None:
                self._quarantine(p, "event.corrupt")
                continue
            if isinstance(data, dict):
                out.append(normalize_event(data))
        out.sort(key=lambda e: e.get("updated_at") or "", reverse=True)
        return out

    def load_event(self, event_id):
        data, err = read_json_file(self.event_path(event_id))
        if err is not None:
            return None, err
        if data is None:
            return None, None
        return normalize_event(data), None

    def save_event(self, event):
        event = normalize_event(event)
        event["updated_at"] = _now_str()
        self.ensure_dirs()
        path = self.event_path(event["id"])
        if path.exists() and event.get("pipeline_ui_version") == "3.9":
            previous, read_error = read_json_file(path)
            if read_error:
                return False, read_error
            if isinstance(previous, dict) and previous.get("pipeline_ui_version") != "3.9":
                try:
                    import shutil
                    backup_dir = self.data_dir / "backups"
                    backup_dir.mkdir(parents=True, exist_ok=True)
                    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                    shutil.copy2(path, backup_dir / (path.stem + ".pre_v3_9_0_" + stamp + ".json"))
                except OSError as error:
                    return False, "旧事件备份失败，未保存：%s" % error
        ok, err = write_json_file(self.event_path(event["id"]), event)
        return ok, err

    def delete_event(self, event_id):
        p = self.event_path(event_id)
        if p.exists():
            try:
                p.unlink()
                return True, None
            except OSError as e:
                return False, str(e)
        return False, "事件不存在"

    # ----- 场景 -----

    def scenario_path(self, scenario_id):
        return self.scenarios_dir / ("%s.json" % sanitize_filename(scenario_id, 80))

    def list_scenarios(self):
        out = []
        if not self.scenarios_dir.exists():
            return out
        for p in sorted(self.scenarios_dir.glob("*.json")):
            data, err = read_json_file(p)
            if err is not None:
                self._quarantine(p, "scenario.corrupt")
                continue
            if isinstance(data, dict):
                out.append(normalize_scenario(data))
        out.sort(key=lambda s: s.get("updated_at") or "", reverse=True)
        return out

    def load_scenario(self, scenario_id):
        data, err = read_json_file(self.scenario_path(scenario_id))
        if err is not None:
            return None, err
        if data is None:
            return None, None
        return normalize_scenario(data), None

    def save_scenario(self, sc):
        sc = normalize_scenario(sc)
        sc["updated_at"] = _now_str()
        self.ensure_dirs()
        ok, err = write_json_file(self.scenario_path(sc["id"]), sc)
        return ok, err

    def delete_scenario(self, scenario_id):
        p = self.scenario_path(scenario_id)
        if p.exists():
            try:
                p.unlink()
                return True, None
            except OSError as e:
                return False, str(e)
        return False, "场景不存在"

    # ----- 执行记录 -----

    def save_run(self, run):
        self.ensure_dirs()
        # ⚠️ id 也要清洗：run id 里若带 : / \ 等 Windows 文件名非法字符，
        #    写文件会失败，而调用方若忽略返回值 → 执行记录静默丢失。
        name = "%s-%s.json" % (sanitize_filename(run.get("event_id", "evt"), 60),
                               sanitize_filename(run.get("id") or _stamp(), 60))
        ok, err = write_json_file(self.runs_dir / name, run)
        return ok, err

    def list_runs(self, event_id=None):
        out = []
        if not self.runs_dir.exists():
            return out
        for p in sorted(self.runs_dir.glob("*.json")):
            data, err = read_json_file(p)
            if err is not None:
                self._quarantine(p, "run.corrupt")
                continue
            if not isinstance(data, dict):
                continue
            if event_id and data.get("event_id") != event_id:
                continue
            out.append(data)
        out.sort(key=lambda r: r.get("started_at") or "", reverse=True)
        return out

    def load_last_run(self, event_id):
        runs = self.list_runs(event_id)
        return runs[0] if runs else None

    # ----- 收件人库（P4：跨事件复用） -----

    def pool_path(self, rid):
        return self.recipients_dir / ("%s.json" % sanitize_filename(rid, 80))

    def list_pool(self):
        """收件人库全部记录，按名称排序。损坏文件改名留证后跳过。"""
        out = []
        if not self.recipients_dir.exists():
            return out
        for p in sorted(self.recipients_dir.glob("*.json")):
            data, err = read_json_file(p)
            if err is not None:
                self._quarantine(p, "recipient.corrupt")
                continue
            if isinstance(data, dict):
                out.append(normalize_pool_recipient(data))
        out.sort(key=lambda r: (r.get("name") or "", r.get("updated_at") or ""))
        return out

    def save_pool_recipient(self, rec):
        """存一条收件人库记录。返回 (ok, error, record)。"""
        rec = normalize_pool_recipient(rec)
        if not rec.get("id"):
            rec["id"] = _new_id("rcp")
        rec["updated_at"] = _now_str()
        self.ensure_dirs()
        ok, err = write_json_file(self.pool_path(rec["id"]), rec)
        return ok, err, rec

    def delete_pool_recipient(self, rid):
        p = self.pool_path(rid)
        if p.exists():
            try:
                p.unlink()
                return True, None
            except OSError as e:
                return False, str(e)
        return False, "收件人不存在"

    def merge_pool_records(self, records):
        """导入库记录并保存新增和同名更新；只补空缺、保留口令、合并标签。"""
        pool = self.list_pool()
        stats = {"added": 0, "updated": 0, "ids": []}
        for record in records:
            src = dict(record)
            src.pop("id", None)
            src["fields"] = {}  # 事件字段不随联系人入库。
            pool, new_rec, is_update = pool_merge_recipient(pool, src)
            target = new_rec
            if is_update:
                target = next((r for r in pool if r.get("name") == src.get("name")), None)
            if target is None:
                continue
            ok, err, saved = self.save_pool_recipient(target)
            if not ok:
                return False, err, stats
            stats["updated" if is_update else "added"] += 1
            if saved["id"] not in stats["ids"]:
                stats["ids"].append(saved["id"])
        return True, None, stats

    # ----- 内部工具 -----

    def _quarantine(self, path, prefix):
        """损坏文件改名留证（不删除、不覆盖）。返回新路径或 None。"""
        p = Path(path)
        try:
            bak = p.with_name("%s-%s%s" % (prefix, _stamp(), p.suffix))
            p.replace(bak)
            return bak
        except OSError:
            return None


# ============================================================================
# v2 → v3 迁移
# ============================================================================

def v2_to_v3_payload(legacy):
    """把 v2 的 lp_data.json 内容转成 v3 的 settings / scenario / event 三件套。

    返回 (settings, scenario, event, notes)。纯函数，便于测试。
    """
    notes = []
    data = legacy if isinstance(legacy, dict) else {}

    settings = dict(DEFAULT_SETTINGS)
    settings["wm_settings"] = dict(data.get("wm_settings") or {})
    settings["gp_owner_mode"] = bool(data.get("gp_owner_mode"))
    settings["gp_owner_pw"] = data.get("gp_owner_pw") or ""
    settings["outdir"] = data.get("outdir") or ""
    settings["filename_tpl"] = data.get("filename_tpl") or DEFAULT_FILENAME_TPL
    settings["auto_update"] = bool(data.get("auto_update", True))
    settings["mail_subject"] = data.get("subject") or ""
    settings["mail_body"] = data.get("body") or ""
    settings["from_addr"] = data.get("from_addr") or ""

    recips = [normalize_recipient(r) for r in (data.get("lps") or [])
              if isinstance(r, dict)]

    # 流水线：v2 有 Word 模板 → 变成一个「生成文档」条目（带转 PDF）
    pipeline = []
    tpl = data.get("word_template") or ""
    if tpl:
        pipeline.append(normalize_item({
            "kind": "generate", "name": "Word 模板",
            "params": {"template_path": tpl, "to_pdf": True},
            "mapping": {},
        }))
        notes.append("Word 模板已转为流水线的一个「生成文档」条目：%s" % tpl)
    else:
        notes.append("v2 未记录原件 PDF 清单（该列表当时不落盘），"
                     "请在流水线里用「源 PDF」步骤重新添加原件")
    if data.get("wt_mappings"):
        notes.append("v2 的外部数据表映射（%d 条占位符）已保留在事件的"
                     " legacy_mappings 里；v2 没有保存那张 Excel 的路径，"
                     "如仍需 from_sheet 请在步骤里重新选择数据表"
                     % len(data.get("wt_mappings") or []))

    mail = {
        "subject": data.get("subject") or "",
        "body": data.get("body") or "",
        "from_addr": data.get("from_addr") or "",
        "cc_default": "", "bcc_default": "",
    }
    options = normalize_options({
        "merge_groups": True,
        "gp_owner_mode": bool(data.get("gp_owner_mode")),
        "gp_owner_pw": data.get("gp_owner_pw") or "",
        "outdir": data.get("outdir") or "",
        "filename_tpl": data.get("filename_tpl") or DEFAULT_FILENAME_TPL,
        "wm_settings": dict(data.get("wm_settings") or {}),
        "constants": {},
    })

    scenario = normalize_scenario({
        "name": "默认场景",
        "description": "由 v2 数据自动迁移生成（可改，不影响已创建的事件）",
        "created_at": _now_str(), "updated_at": _now_str(),
        "recipients": recips, "pipeline": pipeline, "attachments": [],
        "mail": mail, "options": options,
    })

    event = new_event("未命名事件", scenario)
    event["legacy_mappings"] = data.get("wt_mappings") or []
    notes.append("已生成默认场景「默认场景」+ 事件「未命名事件」，"
                 "收件人 %d 位" % len(recips))
    return settings, scenario, event, notes


def migrate_v2_to_v3(store, legacy_path=None):
    """v2 lp_data.json → lp_data/ 目录。返回 (ok, notes, error)。

    规则：原文件改名保留为 lp_data.json.bak，**不删除**。
    """
    legacy = Path(legacy_path) if legacy_path else store.legacy_file
    if not legacy.exists():
        return False, [], "没有找到 %s" % legacy.name
    if store.data_dir.exists() and any(store.data_dir.iterdir()):
        return False, [], "%s 已存在且非空，跳过迁移" % store.data_dir.name

    data, err = read_json_file(legacy)
    if err is not None:
        # 损坏数据绝不覆盖：改名留证并中止
        bak = store._quarantine(legacy, "lp_data.corrupt")
        msg = "%s 读取失败：%s" % (legacy.name, err)
        if bak:
            msg += "（原文件已备份为 %s）" % bak.name
        return False, [], msg
    if not isinstance(data, dict) or not isinstance(data.get("lps"), list):
        bak = store._quarantine(legacy, "lp_data.corrupt")
        msg = "%s 数据结构不合法（缺少 lps 列表）" % legacy.name
        if bak:
            msg += "（原文件已备份为 %s）" % bak.name
        return False, [], msg

    settings, scenario, event, notes = v2_to_v3_payload(data)
    store.ensure_dirs()
    ok, e1 = store.save_settings(settings)
    if not ok:
        return False, notes, "写入 settings.json 失败：%s" % e1
    ok, e2 = store.save_scenario(scenario)
    if not ok:
        return False, notes, "写入场景失败：%s" % e2
    ok, e3 = store.save_event(event)
    if not ok:
        return False, notes, "写入事件失败：%s" % e3

    # 原文件改名保留（不删除），保证可回退
    try:
        legacy.replace(legacy.with_name(legacy.name + ".bak"))
        notes.append("原 %s 已改名保留为 %s.bak（未删除，可随时回退）"
                     % (legacy.name, legacy.name))
    except OSError as e:
        notes.append("原 %s 改名保留失败（数据已迁移，不影响使用）：%s"
                     % (legacy.name, e))
    return True, notes, None


def needs_migration(store):
    """是否需要执行 v2 → v3 迁移。"""
    if store.data_dir.exists() and any(store.data_dir.iterdir()):
        return False
    return store.legacy_file.exists()


def seed_demo(store, example_path=None):
    """首次启动播种示例数据（保留 v2「开箱就有演示数据」的体验）。

    没有 lp_data/、也没有 lp_data.json 时调用。返回 (ok, notes)。
    """
    src = Path(example_path) if example_path else (store.base / "lp_data.example.json")
    data, err = read_json_file(src)
    if not isinstance(data, dict) or not data.get("lps"):
        # 没有示例文件（或文件里没数据）→ 用内置演示数据兜底
        rows, _unknown = parse_import_rows([MOCK_HEADERS] + [list(r) for r in MOCK_LPS])
        data = {"lps": rows,
                "subject": DEFAULT_SUBJECT_V3, "body": DEFAULT_BODY_V3}
    settings, scenario, event, notes = v2_to_v3_payload(data)
    # v3.5.1：场景模板功能已移除；示例数据只保留一个事件
    event["pipeline"] = []
    event["name"] = "示例事件"
    filled = 0
    for rec in list(event["recipients"]):
        if not rec.get("open_pw"):
            rec["open_pw"] = gen_pw()
            filled += 1
        if not rec.get("owner_pw"):
            rec["owner_pw"] = gen_pw()
            filled += 1
    store.ensure_dirs()
    store.save_settings(settings)
    store.save_event(event)

    # P4：示例收件人库（让「收件人库」标签页不是空的，也示范跨事件复用）
    pool_added = 0
    for rec in (data.get("pool") or []):
        if not isinstance(rec, dict) or not rec.get("name"):
            continue
        src = normalize_pool_recipient(rec)
        if not src.get("open_pw"):
            src["open_pw"] = gen_pw()
        if not src.get("owner_pw"):
            src["owner_pw"] = gen_pw()
        ok, err, saved = store.save_pool_recipient(src)
        if not ok:
            notes.append("写入示例收件人库失败：%s" % err)
            break
        pool_added += 1
    if pool_added:
        notes.append("已生成示例收件人库 %d 条（可跨事件复用）" % pool_added)

    notes.append("已生成示例事件「示例事件」，收件人 %d 位%s"
                 % (len(event["recipients"]),
                    ("，自动补齐口令 %d 个" % filled) if filled else ""))
    return True, notes


# ============================================================================
# 流水线：字段解析 / 产出单元 / 步骤执行
# ============================================================================

# 内置字段的别名（用户在模板里可能写成其中任何一种）
BUILTIN_ALIASES = {
    "name": ("name", "client_name", "名称", "收件人名称", "lp名称", "lp_name", "lpname"),
    "wm": ("wm", "水印", "水印名称", "wm_name"),
    "open_pw": ("open_pw", "打开口令", "打开密码", "openpw"),
    "date": ("date", "日期"),
}

# 绝对不能进邮件/文档的字段名（架构保证，不是靠自觉）
FORBIDDEN_FIELD_KEYS = ("owner_pw", "owner_password", "ownerpw",
                         "权限口令", "权限密码", "owner")


def _is_forbidden_key(key):
    k = str(key).strip().strip("{}").strip().lower()
    return k in [x.lower() for x in FORBIDDEN_FIELD_KEYS]


def build_field_map(recipient, event=None, sheet_row=None):
    """返回 {占位符: 值}。

    铁律：权限口令绝不进入映射（键名与值都不进）。
    """
    rec = recipient or {}
    opts = (event or {}).get("options") or {}
    consts = opts.get("constants") or {}
    out = {}
    name = rec.get("name", "")
    base = {
        "name": name,
        "wm": rec.get("wm") or name,
        "open_pw": rec.get("open_pw", ""),
        "date": date.today().strftime("%Y年%m月%d日"),
    }
    for group, keys in BUILTIN_ALIASES.items():
        for k in keys:
            out["{{%s}}" % k] = base[group]
    for k, v in (rec.get("fields") or {}).items():
        if _is_forbidden_key(k):
            continue          # 权限口令类的自定义列一律不映射
        out["{{%s}}" % k] = "" if v is None else str(v)
        out["{{fields.%s}}" % k] = "" if v is None else str(v)
    for k, v in consts.items():
        if _is_forbidden_key(k):
            continue
        out["{{%s}}" % k] = "" if v is None else str(v)
        out["{{constants.%s}}" % k] = "" if v is None else str(v)
    for k, v in (sheet_row or {}).items():
        if _is_forbidden_key(k):
            continue
        out["{{%s}}" % k] = "" if v is None else str(v)
        out["{{sheet.%s}}" % k] = "" if v is None else str(v)
    return out


def scrub_owner_password(text, owner_pw):
    """兜底：把权限口令明文从任何将进入邮件/文档的文本里抹掉。"""
    if not owner_pw:
        return text
    return text.replace(owner_pw, "【权限口令已屏蔽】")


def render_with_fields(tpl, field_map, owner_pw="", escape=False):
    """按字段映射渲染模板（支持 {{字段}} 与 ?字段? 两种写法）。"""
    import html as _h
    out = tpl or ""
    repl = dict(field_map or {})
    for key, val in list(repl.items()):
        if key.startswith("{{") and key.endswith("}}"):
            repl["?" + key[2:-2] + "?"] = val
    for k in sorted(repl, key=len, reverse=True):
        v = repl[k]
        if v is None:
            v = ""
        if escape:
            v = _h.escape(str(v), quote=True)
        out = out.replace(k, str(v))
    return scrub_owner_password(out, owner_pw)


def source_options(event=None, recipient=None):
    """可用的占位符来源（不含外部数据表列）。

    v3.1 起请用 field_source_options()——外部数据表的列只在真的选它时才出现。
    """
    return field_source_options(event, recipient, None)


def apply_mapping_overrides(field_map, mapping_fields, recipient=None, event=None,
                            sheet_row=None):
    """按步骤里显式配置的映射覆盖字段取值（占位符 → 来源键）。"""
    if not mapping_fields:
        return field_map
    rec = recipient or {}
    consts = ((event or {}).get("options") or {}).get("constants") or {}
    fields = rec.get("fields") or {}
    for ph, src in mapping_fields.items():
        if not src or _is_forbidden_key(src):
            continue
        if src == "name":
            val = rec.get("name", "")
        elif src == "wm":
            val = rec.get("wm") or rec.get("name", "")
        elif src == "open_pw":
            val = rec.get("open_pw", "")
        elif src == "date":
            val = date.today().strftime("%Y年%m月%d日")
        elif src.startswith("fields."):
            val = fields.get(src[7:], "")
        elif src.startswith("constants."):
            val = consts.get(src[11:], "")
        elif src.startswith("sheet."):
            val = (sheet_row or {}).get(src[6:], "")
        else:
            val = fields.get(src, consts.get(src, ""))
        field_map[ph] = "" if val is None else str(val)
    return field_map


def placeholder_key(ph):
    """把 {{X}} / ?X? 归一成 X。"""
    s = str(ph or "").strip()
    if s.startswith("{{") and s.endswith("}}"):
        return s[2:-2].strip()
    if s.startswith("?") and s.endswith("?"):
        return s[1:-1].strip()
    return s


def auto_map_placeholder(ph, event=None, recipient=None):
    """按名字自动猜一个来源。返回来源键（猜不到返回空串＝悬空占位符）。"""
    key = placeholder_key(ph)
    if not key:
        return ""
    low = key.lower()
    for group, aliases in BUILTIN_ALIASES.items():
        if low in [a.lower() for a in aliases]:
            return group
    if key in (recipient or {}).get("fields", {}):
        return "fields." + key
    consts = ((event or {}).get("options") or {}).get("constants") or {}
    if key in consts:
        return "constants." + key
    return ""


def item_placeholders(item, event=None, recipient=None, sheet_row=None):
    """条目里已配置的占位符映射：[(占位符, 来源键, 预览值)]。"""
    fields = (item or {}).get("mapping") or {}
    fmap = build_field_map(recipient or {}, event, sheet_row)
    out = []
    for ph, src in fields.items():
        val = fmap.get(ph, "")
        if src and src.startswith("fields."):
            val = (recipient or {}).get("fields", {}).get(src[7:], "")
        elif src and src.startswith("constants."):
            consts = ((event or {}).get("options") or {}).get("constants") or {}
            val = consts.get(src[11:], "")
        elif src and src.startswith("sheet."):
            val = (sheet_row or {}).get(src[6:], "")
        out.append((ph, src, val))
    return out


def field_source_options(event=None, recipient=None, sheet_cols=None):
    """占位符可选来源：[(来源键, 类别, 显示名)]。

    v3.1：来源是**字段级**的，每个占位符自己选一个具体来源；
    外部数据表的列只有在真的要选它时才出现在列表里。
    """
    out = [("Client_Name", "预设", "收件人名称（客户名）"),
           ("name", "内置", "收件人名称"),
           ("wm", "内置", "水印名称"),
           ("open_pw", "内置", "打开口令"),
           ("date", "内置", "日期")]
    for k in sorted((recipient or {}).get("fields", {})):
        out.append(("fields." + k, "本事件字段", k))
    consts = ((event or {}).get("options") or {}).get("constants") or {}
    for k in sorted(consts):
        out.append(("constants." + k, "事件常量", k))
    for k in sorted(sheet_cols or []):
        out.append(("sheet." + k, "外部数据表", k))
    return out


def path_source_options(event=None, recipient=None, sheet_cols=None):
    """附加文件条目可选的路径来源：[(来源键, 类别, 显示名)]。

    与占位符的「字段来源」同一套口径，方便用户理解：
    手填路径 / 收件人清单的附件路径 / 收件人自定义字段 / 外部数据表列。
    """
    out = [(PATH_SOURCE_MANUAL, "固定路径", "手填路径（下面这份清单）"),
           (PATH_SOURCE_RECIPIENT, "收件人", "收件人清单的「附件路径」列")]
    for k in sorted((recipient or {}).get("fields", {})):
        out.append((PATH_SOURCE_FIELD_PREFIX + k, "本事件字段", k))
    for k in sorted(sheet_cols or []):
        out.append((PATH_SOURCE_SHEET_PREFIX + k, "外部数据表", k))
    return out


def path_source_label(src):
    """来源键 → 人话（给提示框/日志用）。"""
    src = str(src or PATH_SOURCE_MANUAL)
    if src == "item.paths":
        return "Excel按LP对应"
    if src == PATH_SOURCE_MANUAL:
        return "手填路径"
    if src == PATH_SOURCE_RECIPIENT:
        return "收件人清单的「附件路径」列"
    if src.startswith(PATH_SOURCE_FIELD_PREFIX):
        return "本事件字段「%s」" % src[len(PATH_SOURCE_FIELD_PREFIX):]
    if src.startswith(PATH_SOURCE_SHEET_PREFIX):
        return "外部数据表「%s」列" % src[len(PATH_SOURCE_SHEET_PREFIX):]
    return src


def resolve_item_paths(item, unit, event, ctx=None):
    """这个附加文件条目在这个产出单元上要附加哪些文件路径。

    返回 (路径列表, 来源说明)。列表里的路径已按 ; 拆开、去空、去重。
    手填路径返回 params["files"]；其余来源按收件人取值，
    取到的值本身也可以再用 ; 分隔多个路径。
    """
    params = (item or {}).get("params") or {}
    src = str(params.get("path_source") or PATH_SOURCE_MANUAL)
    if src == PATH_SOURCE_MANUAL:
        paths = [str(x).strip() for x in (params.get("files") or []) if str(x).strip()]
        return paths, path_source_label(src)

    rec = (unit or {}).get("rep") or {}
    if not rec:
        members = (unit or {}).get("recipients") or []
        rec = members[0] if members else {}
    raw = ""
    if src == "item.paths":
        paths = (params.get("path_map") or {}).get(rec.get("name", ""), [])
        raw = ";".join(paths) if isinstance(paths, list) else str(paths)
    elif src == PATH_SOURCE_RECIPIENT:
        raw = ";".join(str(p) for p in (rec.get("attachments") or []))
    elif src.startswith(PATH_SOURCE_FIELD_PREFIX):
        raw = str((rec.get("fields") or {}).get(src[len(PATH_SOURCE_FIELD_PREFIX):], "") or "")
    elif src.startswith(PATH_SOURCE_SHEET_PREFIX):
        col = src[len(PATH_SOURCE_SHEET_PREFIX):]
        row = None
        if ctx is not None:
            row = ctx.sheet_for(item, rec)
        raw = str((row or {}).get(col, "") or "")
    paths = split_paths(raw)
    out, seen = [], set()
    for p in paths:
        k = p.lower()
        if k not in seen:
            seen.add(k)
            out.append(p)
    return out, path_source_label(src)


def group_key_of(recipient):
    """合并组 key：没有组的收件人自成一组（§+名字，不会和真实组名撞）。"""
    g = (recipient.get("group") or "").strip()
    return g if g else "§" + (recipient.get("name") or "")


def build_units(event, mode):
    """按输出粒度列出产出单元。

    返回 [(unit_key, info)]；info 含 recipients / rep / group_key / lp_name。
    """
    recips = event.get("recipients") or []
    if mode == "all":
        return [("__all__", {"recipients": recips, "rep": None,
                             "group_key": "__all__", "lp_name": ""})]
    if mode == "per_group":
        groups, order = {}, []
        for r in recips:
            key = group_key_of(r)
            if key not in groups:
                groups[key] = []
                order.append(key)
            groups[key].append(r)
        out = []
        for key in order:
            members = groups[key]
            out.append((key, {"recipients": members, "rep": members[0],
                              "group_key": key,
                              "lp_name": members[0].get("name", ""),
                              "group": "" if key.startswith("§") else key}))
        return out
    out = []
    for r in recips:
        out.append((r.get("name") or "", {"recipients": [r], "rep": r,
                                          "group_key": group_key_of(r),
                                          "lp_name": r.get("name", ""),
                                          "group": (r.get("group") or "").strip()}))
    return out


def item_output_name(event, unit, source_path, orig_name=None, filename_tpl=None):
    """产物文件名（不含扩展名）。

    orig_name: 覆盖「原件名」占位符的取值。**生成文档条目用它自己的条目名**
    （否则两个条目用同一个 Word 模板会产出同名文件，附件包里就撞名了）；
    附加文件条目用源文件名（原件名本来就该是那个文件）。
    filename_tpl: v3.5.1 起支持**条目级**模板（Word-MailMerge 参数区填写）；
    为空时回退旧的事件级 options.filename_tpl（兼容旧数据）。
    """
    opts = event.get("options") or {}
    tpl = (str(filename_tpl).strip() if filename_tpl else "") or \
        opts.get("filename_tpl") or DEFAULT_FILENAME_TPL
    rec = unit.get("rep") or {}
    fields = build_field_map(rec, event)
    stem = Path(source_path).stem if source_path else ""
    if orig_name is not None and str(orig_name).strip():
        stem = sanitize_filename(str(orig_name).strip(), 60)
    mapping = {
        "{{原件名}}": stem,
        "{{LP名称}}": sanitize_filename(rec.get("name", ""), 60),
        "{{收件人名称}}": sanitize_filename(rec.get("name", ""), 60),
        "{{水印名称}}": sanitize_filename(rec.get("wm") or rec.get("name", ""), 60),
        "{{日期}}": date.today().strftime("%Y%m%d"),
    }
    out = tpl
    for k, v in mapping.items():
        out = out.replace(k, v)
    for k, v in fields.items():
        if k.startswith("{{") and k.endswith("}}"):
            inner = k[2:-2]
            if inner in ("name", "名称", "收件人名称", "lp名称", "lp_name",
                         "lpname", "wm", "水印", "水印名称", "wm_name"):
                continue      # 已经用带清洗的版本替换过了
            out = out.replace(k, sanitize_filename(str(v), 40))
    out = sanitize_filename(out.strip(), 90)
    return out or "output"


def load_sheet_data(path, key_col=""):
    """读外部 Excel 数据表。返回 (sheet_map, warnings)。

    sheet_map = {key 列的值: {列名: 值}}
    """
    if not path or not str(path).strip():
        return {}, ["from_sheet 步骤没有选择数据表"]
    p = Path(str(path).strip())
    if not p.exists():
        return {}, ["外部数据表不存在：%s" % p]
    try:
        rows = sheet_to_dicts(read_spreadsheet(p))
    except Exception as e:
        return {}, ["读取外部数据表失败：%s" % e]
    if not rows:
        return {}, ["外部数据表是空的：%s" % p]
    cols = list(rows[0].keys())
    if not key_col:
        key_col = cols[0]
    if key_col not in cols:
        return {}, ["外部数据表里没有匹配列「%s」（现有列：%s）"
                   % (key_col, "、".join(cols))]
    out = {}
    for r in rows:
        k = str(r.get(key_col, "")).strip()
        if k:
            out[k] = r
    return out, []


# ---------------------------------------------------------------------------
# Word 实例复用（v3.1）
# ---------------------------------------------------------------------------
# 坑（实测）：WordBatch 用 DispatchEx 起**私有** Word 进程，close() 只释放引用、
# 不调 Quit（Quit 之后再释放 COM 引用会触发 0x800706be 致命异常，直接杀进程）。
# 于是「每次生成新建一个 WordBatch」= 每次生成残留一个 WINWORD.EXE（每个约
# 170MB）；实测一个会话能堆到 71 个，机器被拖垮（空闲内存 7.6GB → 4.1GB）。
# 所以按线程缓存一个共享实例：一个进程最多只留一个 Word 进程。
_WORD_LOCAL = threading.local()


def default_word_batch():
    """当前线程共享的 Word 实例（整个进程最多一个 WINWORD）。"""
    wb = getattr(_WORD_LOCAL, "wb", None)
    if wb is None:
        wb = wm.WordBatch()
        _WORD_LOCAL.wb = wb
    return wb


class RunContext:
    """一次流水线执行期的上下文。"""

    def __init__(self, event, outdir, word=None):
        self.event = event
        self.outdir = Path(outdir)
        self.word = word
        self._sheet_cache = {}
        self.warnings = []
    def sheet_for(self, item, recipient):
        """收件人对应的外部数据行。

        只有这个条目的字段映射里用到了 sheet. 来源，才会去读那张外部表；
        每个条目可以配自己的表，按 (路径, 匹配列) 缓存，避免重复读。
        """
        if not recipient or not item_uses_sheet(item):
            return None
        sh = item.get("sheet") or {}
        path = str(sh.get("path") or "").strip()
        if not path:
            return None
        key = (path, sh.get("key_col", ""))
        if key not in self._sheet_cache:
            m, w = load_sheet_data(path, sh.get("key_col", ""))
            self._sheet_cache[key] = (m, w)
            self.warnings += w
        m, _w = self._sheet_cache[key]
        match_field = sh.get("match_col") or "name"
        if match_field == "name":
            val = recipient.get("name", "")
        else:
            val = (recipient.get("fields") or {}).get(match_field, "")
        return m.get(str(val).strip())

    def warn(self, msg):
        self.warnings.append(msg)


def _wm_text_for(unit, event):
    """本单元的水印文字。"""
    rep = unit.get("rep")
    if rep:
        return rep.get("wm") or rep.get("name") or ""
    recips = unit.get("recipients") or []
    if recips:
        return recips[0].get("wm") or recips[0].get("name") or ""
    return event.get("name", "")


def _passwords_for(unit, event):
    """本单元的 (打开口令, 权限口令)。GP 统一口令优先。"""
    opts = event.get("options") or {}
    rep = unit.get("rep") or {}
    if not rep:
        # 全员共用一份（rep 为空）→ 取第一位收件人的口令
        recips = unit.get("recipients") or []
        rep = recips[0] if recips else {}
    open_pw = rep.get("open_pw", "")
    owner_pw = rep.get("owner_pw", "")
    if opts.get("gp_owner_mode") and opts.get("gp_owner_pw"):
        owner_pw = opts["gp_owner_pw"]
    return open_pw, owner_pw


def _permission_mask(perms):
    """权限勾选 → PyMuPDF 权限位掩码。

    默认（未勾任何项）= 仅打印，与既有铁律一致。
    """
    perms = perms if isinstance(perms, (list, tuple)) else DEFAULT_PERMISSIONS
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


def execute_item(item, unit, event, ctx):
    """执行一个条目在某个产出单元上的处理。

    kind=generate：Word 模板 → .docx →（可选）PDF →（可选）水印 →（可选）加密
    kind=attach  ：磁盘文件 →（可选，仅 PDF）水印 →（可选，仅 PDF）加密

    返回 (产物文件列表, 警告列表)。失败抛异常。
    """
    import shutil
    kind = item.get("kind") or "generate"
    params = item.get("params") or {}
    item_id = item.get("id") or "item"
    outdir = ctx.outdir / sanitize_filename(item_id, 40)
    outdir.mkdir(parents=True, exist_ok=True)
    warnings = []
    opts = event.get("options") or {}

    if kind == "generate":
        tpl = str(params.get("template_path", "")).strip()
        if not tpl:
            raise RuntimeError("「生成文档」没有选择 Word 模板")
        if not Path(tpl).exists():
            raise RuntimeError("Word 模板不存在：%s" % tpl)
        rep = unit.get("rep")
        # 只有映射里用到了 sheet. 来源，才去查外部数据表行
        sheet_row = ctx.sheet_for(item, rep)
        fmap = build_field_map(rep or {}, event, sheet_row)
        apply_mapping_overrides(fmap, item.get("mapping") or {},
                                rep or {}, event, sheet_row)
        base = item_output_name(
            event, unit, tpl,
            orig_name=(item.get("name") or "").strip() or None,
            filename_tpl=(params.get("filename_tpl") or ""))
        docx = outdir / (base + ".docx")
        wm.generate_notice_docx(tpl, str(docx), fmap, batch=ctx.word)
        cur = docx
    else:
        files, src_label = resolve_item_paths(item, unit, event, ctx)
        per_recipient = str(params.get("path_source") or PATH_SOURCE_MANUAL) != PATH_SOURCE_MANUAL
        if not files:
            if per_recipient:
                # 按收件人取路径时，某个人没填是正常的：跳过并提醒，不算失败
                who = (unit.get("rep") or {}).get("name", "") or unit.get("lp_name", "")
                warnings.append("%s：按「%s」没取到文件路径，已跳过"
                                % (who or "该收件人", src_label))
                return [], warnings
            raise RuntimeError("「附加文件」没有选择任何文件")
        if per_recipient and not (unit.get("rep") or {}):
            warnings.append("产出粒度是「全员共用一份」，按收件人取路径将以第一位收件人的路径为准")
        cur = []
        for src_path in files:
            sp = str(src_path).strip()
            if not sp:
                continue
            if not Path(sp).exists():
                if per_recipient:
                    warnings.append("文件不存在，已跳过：%s" % sp)
                    continue
                raise RuntimeError("文件不存在：%s" % sp)
            base = item_output_name(event, unit, sp)
            if params.get("preserve_name"):
                base = Path(sp).stem
            dst = outdir / (base + Path(sp).suffix)
            suffix_index = 2
            while any(str(existing).lower() == str(dst).lower() for existing in cur):
                dst = outdir / (base + "_%d" % suffix_index + Path(sp).suffix)
                suffix_index += 1
            shutil.copy2(sp, str(dst))
            cur.append(dst)
        if not cur:
            if per_recipient:
                return [], warnings
            raise RuntimeError("「附加文件」没有选择任何文件")

    # ---- 处理链：转 PDF → 加水印 → 加密（顺序固定） ----
    if params.get("to_pdf"):
        converted = []
        reserved = {p.name.lower() for p in ([cur] if isinstance(cur, Path) else cur)
                    if p.suffix.lower() == ".pdf"}
        for source in [cur] if isinstance(cur, Path) else list(cur):
            if source.suffix.lower() == ".pdf":
                converted.append(source)
            else:
                target = outdir / (source.stem + ".pdf")
                suffix_index = 2
                while target.name.lower() in reserved:
                    target = outdir / (source.stem + "_%d.pdf" % suffix_index)
                    suffix_index += 1
                reserved.add(target.name.lower())
                wm.file_to_pdf(str(source), str(target), batch=ctx.word)
                converted.append(target)
        cur = converted[0] if len(converted) == 1 else converted

    if params.get("watermark"):
        targets = [cur] if isinstance(cur, Path) else list(cur)
        new_targets = []
        for src in targets:
            if Path(src).suffix.lower() != ".pdf":
                # 非 PDF 不能加水印：原样保留（并提醒）
                warnings.append("%s 不是 PDF，已跳过水印" % Path(src).name)
                new_targets.append(src)
                continue
            text = _wm_text_for(unit, event)
            if not (text or "").strip():
                raise RuntimeError("水印文字为空（收件人的水印名称与名称都为空）")
            out = outdir / (Path(src).stem + ".pdf")
            info = wm.watermark_pdf(str(src), str(out), text, opts.get("wm_settings"))
            warnings += info.get("warnings", [])
            new_targets.append(out)
        cur = new_targets[0] if len(new_targets) == 1 else new_targets

    if params.get("encrypt"):
        targets = [cur] if isinstance(cur, Path) else list(cur)
        new_targets = []
        open_pw, owner_pw = _passwords_for(unit, event)
        for src in targets:
            if Path(src).suffix.lower() != ".pdf":
                warnings.append("%s 不是 PDF，已跳过加密" % Path(src).name)
                new_targets.append(src)
                continue
            if not open_pw:
                raise RuntimeError("收件人没有打开口令")
            if not owner_pw:
                raise RuntimeError("收件人没有权限口令")
            out = outdir / (Path(src).stem + ".pdf")
            info = wm.encrypt_pdf(str(src), str(out), open_pw, owner_pw,
                                  permissions=params.get("permissions"))
            warnings += info.get("warnings", [])
            new_targets.append(out)
        cur = new_targets[0] if len(new_targets) == 1 else new_targets

    if isinstance(cur, Path):
        return [cur], warnings
    return list(cur), warnings


def run_items(event, outdir, word=None, only_recips=None, log=None, progress=None):
    """执行整个流水线，返回 (outputs, warnings, errors)。

    outputs: {step_id: {unit_key: [文件路径]}}
    errors:  [(unit_key, step_id, 错误信息)]
    only_recips: 只跑这些收件人（按收件人名过滤，「只重跑失败项」用）
    """
    def _log(msg):
        if log:
            log(msg)

    opts = event.get("options") or {}
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    items = [i for i in (event.get("pipeline") or []) if i.get("enabled", True)]
    outputs = {}
    errors = []
    warnings = []
    if not items:
        return outputs, warnings, [("", "", "流水线为空，没有任何启用的条目")]

    ctx = RunContext(event, outdir, word=word)
    total = len(items)
    for si, it in enumerate(items):
        iid = it.get("id") or ("item%d" % si)
        gran = it.get("granularity") or "per_lp"
        units = build_units(event, gran)
        cur = {}
        for ukey, unit in units:
            target_scope = (it.get("params") or {}).get("target_scope")
            if target_scope and not _attachment_matches(target_scope, unit):
                continue
            if only_recips is not None:
                names = set(r.get("name", "") for r in unit.get("recipients") or [])
                if not (names & set(only_recips)):
                    continue
            try:
                files, warns = execute_item(it, unit, event, ctx)
                cur[ukey] = files
                warnings += warns
                _log("[条目%d] %s · %s → %s"
                     % (si + 1, it.get("name") or ITEM_KIND_LABELS.get(it.get("kind")),
                        unit.get("lp_name") or unit.get("group_key"),
                        files[0].name if files else ""))
            except Exception as e:
                errors.append((ukey, iid, str(e)))
                _log("[条目%d 失败] %s：%s"
                     % (si + 1, unit.get("lp_name") or unit.get("group_key"), e))
        outputs[iid] = cur
        if progress:
            progress(int((si + 1) * 100 / total))
    warnings += ctx.warnings
    return outputs, warnings, errors


# ============================================================================
# 执行引擎：预检 / 附件组装 / 草稿 / 执行记录
# ============================================================================

def _attachment_matches(att, unit):
    """静态附件是否属于这个收件人/合并组。"""
    scope = att.get("scope", "all")
    if scope == "all":
        return True
    names = set(r.get("name", "") for r in unit.get("recipients") or [])
    if scope == "specific":
        targets = set(att.get("targets") or [])
        return bool(names & targets)
    if scope == "group":
        g = (att.get("group") or "").strip()
        return bool(g) and g == (unit.get("group") or "").strip()
    return False


def event_uses_recipient_paths(event):
    """流水线上是否已有条目按「收件人清单的附件路径」取路径。

    有的话，收件人页那一列路径已经由条目产出（会按产出粒度重命名/加密），
    就不再原样重复附加一次，避免收件人收到两份同样的文件。
    """
    for it in (event or {}).get("pipeline") or []:
        if not it.get("enabled", True):
            continue
        if (it.get("kind") or "generate") != "attach":
            continue
        src = str(((it.get("params") or {}).get("path_source") or PATH_SOURCE_MANUAL))
        if src in (PATH_SOURCE_RECIPIENT, "item.paths"):
            return True
    return False


def _file_digest(path):
    """文件内容指纹（大小 + SHA256）。读不了返回 None（不参与内容去重）。"""
    import hashlib
    try:
        p = Path(path)
        h = hashlib.sha256()
        with p.open("rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return "%d:%s" % (p.stat().st_size, h.hexdigest())
    except OSError:
        return None


def assemble_attachments(event, unit, final_outputs, dup_log=None, complete=False):
    """组装附件包。返回 (文件列表, 缺失文件列表)。

    附件包 = 所有流水线条目的产物 + 事件级静态附件 + 收件人专属附件。
    （每个条目返回的已经是最终产物——docx 已转 pdf、pdf 已加水印/加密。）
    去重两层：先按路径（大小写不敏感），再按内容指纹（大小+SHA256）——
    合并组内同内容不同文件名的副本只保留第一份。
    dup_log 传入 list 时，每丢弃一个重复附件记 (保留的, 被丢弃的)。
    """
    files, missing = [], []

    def _add(p):
        p = str(p or "").strip()
        if not p:
            return
        if Path(p).exists():
            files.append(Path(p))
        else:
            missing.append(p)

    for p in final_outputs or []:
        _add(p)
    for a in ([] if complete else event.get("attachments") or []):
        if _attachment_matches(a, unit):
            _add(a.get("path"))
    for r in ([] if complete else unit.get("recipients") or []):
        for p in r.get("attachments") or []:
            if event_uses_recipient_paths(event):
                break          # 已由流水线条目按路径来源产出，不重复附加
            _add(p)
    # 去重：仅按路径（同一路径只挂一份）。
    # 内容指纹相同但路径不同的文件【全部保留、不做去重】——只在 dup_log 里
    # 记 (首份, 重复份) 供调用方提示（用户拍板 2026-10-24：提示而非去重）。
    seen_path, seen_hash, out = set(), {}, []
    for p in files:
        key = str(p).lower()
        if key in seen_path:
            continue                        # 同一路径 = 同一个文件，只挂一份
        seen_path.add(key)
        digest = _file_digest(p)
        if digest is not None and digest in seen_hash:
            if dup_log is not None:
                dup_log.append((seen_hash[digest], p))   # 仅记录，不丢弃
        else:
            if digest is not None:
                seen_hash[digest] = p
        out.append(p)                       # 无论指纹是否重复都保留
    return out, missing


def attachment_fingerprint_warnings(event):
    """按草稿单元（=每封邮件）扫描附件的内容指纹重复。

    范围：事件级静态附件 + 成员的收件人专属附件（不含流水线产物——
    每人一份的产物按人各有水印/口令，内容天然不同；all/组粒度产物是单文件）。
    只报告"同封邮件内、不同路径、指纹（大小+SHA256）相同"的组合，
    **不改任何数据、不去重**（用户拍板 2026-10-24）。
    返回 [str] 警告文案。
    """
    warns = []
    for key, unit in draft_units(event):
        paths = []
        for a in event.get("attachments") or []:
            if _attachment_matches(a, unit):
                pth = str(a.get("path") or "").strip()
                if pth and Path(pth).exists():
                    paths.append(pth)
        if not event_uses_recipient_paths(event):
            for r in unit.get("recipients") or []:
                for pth in r.get("attachments") or []:
                    if pth and Path(pth).exists():
                        paths.append(str(pth))
        by_digest = {}
        for pth in paths:
            d = _file_digest(pth)
            if d is None:
                continue
            prev = by_digest.get(d)
            if prev is None:
                by_digest[d] = pth
            elif prev.lower() != pth.lower():
                warns.append("[附件重复提示] %s 与 %s 内容完全相同（不同路径），"
                             "两份都会保留，请人工确认" % (Path(prev).name, Path(pth).name))
    return warns


def draft_units(event):
    """草稿单元：按事件的发送模式决定「几封邮件」。

    per_lp    每人一份 → 每位收件人一封
    per_group 按组合并 → 同合并组一封（没有组的自成一组）
    all       全部合并 → 所有人一封

    返回 [(unit_key, unit)]。
    """
    mode = send_mode_of(event)
    recips = event.get("recipients") or []
    if mode == "all":
        if not recips:
            return []
        return [("§all", {
            "recipients": list(recips), "rep": recips[0],
            "group_key": "§all",
            "lp_name": recips[0].get("name", ""),
            "group": ""})]
    if mode == "per_lp":
        out = []
        for r in recips:
            out.append((r.get("name", ""), {
                "recipients": [r], "rep": r,
                "group_key": group_key_of(r),
                "lp_name": r.get("name", ""),
                "group": (r.get("group") or "").strip()}))
        return out
    groups, order = {}, []
    for r in recips:
        key = group_key_of(r)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)
    out = []
    for key in order:
        members = groups[key]
        out.append((key, {
            "recipients": members, "rep": members[0],
            "group_key": key,
            "lp_name": members[0].get("name", ""),
            "group": "" if key.startswith("§") else key}))
    return out


def _merge_addresses(members, key, extra=""):
    """合并组内收件人的 To/CC/BCC（去重、保持顺序）。"""
    out = []
    for m in members:
        for addr in split_mails(m.get(key, "") or ""):
            if addr not in out:
                out.append(addr)
    for addr in split_mails(extra or ""):
        if addr not in out:
            out.append(addr)
    return out


def create_outlook_draft(rec, attachments, subject_tpl, body_tpl, field_map=None,
                          owner_pw="", extra_cc="", extra_bcc=""):
    """在 Outlook 草稿箱创建邮件（只 Save，绝不 Send）。

    铁律：正文/标题里的权限口令会被 scrub 掉（架构 + 兜底双保险）。
    """
    import win32com.client as win32
    ol = win32.Dispatch("Outlook.Application")
    mail = ol.CreateItem(0)  # 0 = olMailItem
    to = _merge_addresses([rec], "to")
    cc = _merge_addresses([rec], "cc", extra_cc)
    bcc = _merge_addresses([rec], "bcc", extra_bcc)
    if to:
        mail.To = ";".join(to)
    if cc:
        mail.CC = ";".join(cc)
    if bcc:
        mail.BCC = ";".join(bcc)
    mail.Subject = render_with_fields(subject_tpl, field_map, owner_pw)
    html = render_with_fields(body_tpl, field_map, owner_pw, escape=True)
    # v3.7.0：编辑器插入的图片是 data URI；Outlook 桌面版不显示 data URI，
    # 这里转成 cid 内嵌附件（附件先加，再给 HTMLBody 用 cid: 引用）。
    temp_imgs = []
    if "data:image/" in html:
        import base64
        import re as _re
        import tempfile

        def _inline(m):
            mime = (m.group(1) or "png").lower()
            b64 = m.group(2)
            cid = "docmailer_img_%d@local" % len(temp_imgs)
            try:
                raw = base64.b64decode(b64)
            except Exception:
                return m.group(0)
            fd = tempfile.NamedTemporaryFile(delete=False, suffix="." + mime)
            fd.write(raw)
            fd.close()
            temp_imgs.append(fd.name)
            att = mail.Attachments.Add(fd.name)
            try:
                att.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
            except Exception:
                pass
            return "cid:" + cid

        html = _re.sub(r'data:image/([a-zA-Z0-9+.-]+);base64,([A-Za-z0-9+/=\s]+)',
                       _inline, html)
    mail.HTMLBody = html
    from_addr = rec.get("from_addr", "")
    if from_addr:
        mail.SentOnBehalfOfName = from_addr
    for f in attachments:
        mail.Attachments.Add(str(Path(f).resolve()))  # Outlook 是独立进程，必须绝对路径
    try:
        mail.Save()  # 只存草稿；不调用 Send
    finally:
        # 草稿保存后内嵌附件已复制进邮件，临时图片文件可以清理
        for _p in temp_imgs:
            try:
                Path(_p).unlink(missing_ok=True)
            except Exception:
                pass
    return mail


def preflight(event, store=None):
    """预检。返回 (errors, warnings, info)。

    errors 会中止执行；warnings 只提示（缺失附件等会被跳过）。
    """
    errors, warnings = [], []
    recips = event.get("recipients") or []
    opts = event.get("options") or {}
    steps = [s for s in (event.get("pipeline") or []) if s.get("enabled", True)]

    # --- 收件人 ---
    if not recips:
        errors.append("收件人清单为空")
    seen_pw = {}
    for r in recips:
        tag = (r.get("name") or "?")[:18]
        if not (r.get("name") or "").strip():
            errors.append("%s：收件人名称为空" % tag)
        if not r.get("to"):
            warnings.append("%s：没有邮箱，将只生成文件、不建草稿" % tag)
        else:
            for m in split_mails(r.get("to", "")):
                if "@" not in m or "." not in m.split("@")[-1]:
                    warnings.append("%s：邮箱格式可疑：%s" % (tag, m))
        if not r.get("open_pw"):
            errors.append("%s：打开口令为空" % tag)
        gp_on = bool(opts.get("gp_owner_mode")) and bool(opts.get("gp_owner_pw"))
        if not r.get("owner_pw") and not gp_on:
            errors.append("%s：权限口令为空（留档缺失将无法解除限制）" % tag)
        elif r.get("open_pw") and r.get("open_pw") == r.get("owner_pw"):
            errors.append("%s：权限口令与打开口令相同（收件人将获得全部权限）" % tag)
        if r.get("open_pw"):
            seen_pw.setdefault(r["open_pw"], []).append(tag)
    for _pw, names in seen_pw.items():
        if len(names) > 1:
            errors.append("打开口令重复（%s）：同一口令可打开对方文件" % "、".join(names))

    # --- GP 统一口令冲突 ---
    gp = opts.get("gp_owner_pw") if opts.get("gp_owner_mode") else ""
    if gp:
        clash = [r.get("name", "") for r in recips if r.get("open_pw") == gp]
        if clash:
            errors.append("GP 统一权限口令与以下收件人的打开口令相同（对方将获得全部权限）：%s"
                          % "、".join(clash[:8]))

    # --- 合并组 ---
    groups = {}
    for r in recips:
        g = (r.get("group") or "").strip()
        if g:
            groups.setdefault(g, []).append(r)
    for g, members in groups.items():
        if len(set(m.get("open_pw", "") for m in members)) > 1:
            warnings.append("合并组「%s」内打开口令不同，正文只会写第一个人的口令，"
                            "建议统一口令或拆开发送" % g)

    # --- 附件存在性 ---
    for a in event.get("attachments") or []:
        p = (a.get("path") or "").strip()
        if p and not Path(p).exists():
            warnings.append("附件不存在（将被跳过）：%s" % p)
    for r in recips:
        for p in r.get("attachments") or []:
            if not Path(p).exists():
                warnings.append("%s：专属附件不存在（将被跳过）：%s"
                                % (r.get("name", ""), p))

    # Saved per-person file lists are final attachment selections.
    custom = event.get("process_overrides") or {}
    automatic = [r for r in recips if r.get("name", "") not in custom]
    for name, paths in custom.items():
        paths = split_paths(paths) if isinstance(paths, str) else paths
        for path in paths if isinstance(paths, list) else []:
            if not Path(path).is_file():
                errors.append("%s：保存的自定义文件不存在：%s" % (name, path))
    # --- 流水线条目 ---
    if not steps and automatic:
        errors.append("流水线为空，没有任何启用的条目")
    for i, it in enumerate(steps if automatic else []):
        label = "第%d条 %s" % (i + 1, it.get("name") or
                               ITEM_KIND_LABELS.get(it.get("kind"), ""))
        kind = it.get("kind") or "generate"
        params = it.get("params") or {}
        gran = it.get("granularity") or "per_lp"
        if kind == "generate":
            tp = (params.get("template_path") or "").strip()
            if not tp:
                errors.append("%s：没有选择 Word 模板" % label)
            elif not Path(tp).exists():
                errors.append("%s：Word 模板不存在：%s" % (label, tp))
            else:
                try:
                    phs = wm.detect_placeholders(tp)
                except Exception as e:
                    warnings.append("%s：读取模板占位符失败：%s" % (label, e))
                    phs = []
                mapped = it.get("mapping") or {}
                rec0 = recips[0] if recips else {}
                for ph in phs:
                    src = mapped.get(ph) or auto_map_placeholder(ph, event, rec0)
                    if not src:
                        errors.append("%s：占位符 %s 没有对应字段（悬空占位符）" % (label, ph))
                    elif str(src).startswith("sheet."):
                        sh = it.get("sheet") or {}
                        sp = (sh.get("path") or "").strip()
                        if not sp:
                            errors.append("%s：占位符 %s 用了外部数据表，但还没选表" % (label, ph))
                        elif not Path(sp).exists():
                            errors.append("%s：外部数据表不存在：%s" % (label, sp))
        else:
            psrc = str(params.get("path_source") or PATH_SOURCE_MANUAL)
            if psrc != PATH_SOURCE_MANUAL:
                # v3.5.0：「从 Excel 导入路径」后按收件人取路径，手填清单允许为空
                anyone = False
                for _r in recips:
                    _unit = {"rep": _r, "recipients": [_r],
                             "lp_name": _r.get("name", ""),
                             "group_key": group_key_of(_r),
                             "group": (_r.get("group") or "").strip()}
                    try:
                        _paths, _lab = resolve_item_paths(item, _unit, event)
                    except Exception as _e:
                        warnings.append("%s：%s 取路径失败：%s"
                                        % (label, _r.get("name", ""), _e))
                        continue
                    if _paths:
                        anyone = True
                    for sp in _paths:
                        if not Path(str(sp)).exists():
                            warnings.append("%s：%s 的附件不存在（将跳过）：%s"
                                            % (label, _r.get("name", ""), sp))
                if not anyone:
                    warnings.append("%s：按「%s」没有取到任何路径，处理时会被跳过"
                                    % (label, path_source_label(psrc)))
            else:
                files = params.get("files") or []
                if not files:
                    errors.append("%s：没有选择任何文件" % label)
                for sp in files:
                    if not Path(str(sp)).exists():
                        errors.append("%s：文件不存在：%s" % (label, sp))
        if kind == "attach":
            paths = list(params.get("files") or [])
            if params.get("path_source", PATH_SOURCE_MANUAL) != PATH_SOURCE_MANUAL:
                paths = []
                for recipient in automatic:
                    selected, _ = resolve_item_paths(it, {"rep": recipient}, event)
                    paths.extend(selected)
            for source in paths:
                suffix = Path(source).suffix.lower()
                if params.get("input_type") == "pdf" and suffix != ".pdf":
                    errors.append("%s：PDF入口只能添加PDF文件：%s" % (label, source))
                if params.get("to_pdf") and suffix != ".pdf" and suffix not in wm.PDF_CONVERT_EXTENSIONS:
                    errors.append("%s：该格式不支持转PDF：%s" % (label, source))
                if params.get("input_type") in ("file", "excel") and suffix != ".pdf" and not params.get("to_pdf") and (params.get("watermark") or params.get("encrypt")):
                    errors.append("%s：非PDF文件需先勾选转PDF，才能加水印/加密" % label)
        if params.get("encrypt"):
            if not params.get("watermark") and kind == "generate" \
                    and not params.get("to_pdf"):
                warnings.append("%s：只勾了加密，但产物是 .docx（加密只对 PDF 生效）"
                                % label)
        if gran == "all":
            warnings.append("%s：全员共用一份，文件口令取第一位收件人的口令" % label)
        if kind == "attach" and params.get("watermark"):
            for sp in (params.get("files") or []):
                if Path(str(sp)).suffix.lower() != ".pdf":
                    warnings.append("%s：%s 不是 PDF，水印不会生效" % (label, Path(str(sp)).name))

    info = {
        "recipients": len(recips),
        "steps": len(steps),
        "attachments": len(event.get("attachments") or []),
        "drafts": len(draft_units(event)),
        "expected_files": len(steps) * len(draft_units(event)),
    }
    warnings.extend(attachment_fingerprint_warnings(event))
    return errors, warnings, info


def _item_units_for(item, unit):
    """这个条目在某个草稿单元里对应的「产出单位」列表。

    条目粒度与发送模式粒度可能不同，要桥接：
      条目 per_lp    → 单元内每位收件人各一份
      条目 per_group → 按收件人所在组各一份（组内第一个收件人当代表）
      条目 all      → 全员共用一份
    """
    gran = (item or {}).get("granularity") or "per_lp"
    recips = unit.get("recipients") or []
    if gran == "all":
        return [unit]
    if gran == "per_group":
        seen, out = set(), []
        for r in recips:
            k = group_key_of(r)
            if k in seen:
                continue
            seen.add(k)
            one = dict(unit)
            one["rep"] = r
            one["group_key"] = k
            one["lp_name"] = r.get("name", "")
            out.append(one)
        return out
    out = []
    for r in recips:
        one = dict(unit)
        one["rep"] = r
        one["lp_name"] = r.get("name", "")
        one["group_key"] = group_key_of(r)
        out.append(one)
    return out


def _item_output_key(item, one):
    """产出单位在 outputs 字典里的 key（必须与 build_units 的 key 一致）。"""
    gran = (item or {}).get("granularity") or "per_lp"
    if gran == "all":
        return "__all__"
    if gran == "per_group":
        return one.get("group_key") or ""
    return one.get("lp_name") or ""


def _item_preview_files(event, item, unit):
    """这个条目在这个草稿单元里会产出的文件名（预览用，不执行）。"""
    out = []
    params = (item or {}).get("params") or {}
    kind = (item or {}).get("kind") or "generate"
    for one in _item_units_for(item, unit):
        if kind == "attach":
            paths, _ = resolve_item_paths(item, one, event)
            for sp in paths:
                base = item_output_name(event, one, str(sp))
                extension = ".pdf" if params.get("to_pdf") else Path(str(sp)).suffix
                out.append(base + extension)
        else:
            src = params.get("template_path", "")
            base = item_output_name(
                event, one, src,
                orig_name=(item.get("name") or "").strip() or None,
                filename_tpl=(params.get("filename_tpl") or ""))
            out.append(base + (".pdf" if params.get("to_pdf") else ".docx"))
    return out


def processing_signature(event):
    """Fingerprint processing settings and source-file versions, excluding mail edits."""
    import hashlib
    paths = set()
    for item in event.get("pipeline") or []:
        params = item.get("params") or {}
        paths.update(str(p) for p in params.get("files") or [])
        paths.update(p for values in (params.get("path_map") or {}).values() for p in values)
        paths.add(str(params.get("template_path") or ""))
        paths.add(str((item.get("sheet") or {}).get("path") or ""))
    for recipient in event.get("recipients") or []:
        paths.update(recipient.get("attachments") or [])
    for item in event.get("pipeline") or []:
        if item.get("kind") == "attach":
            for _, unit in build_units(event, item.get("granularity") or "per_lp"):
                selected, _ = resolve_item_paths(item, unit, event)
                paths.update(selected)
    versions = []
    for path in sorted(paths - {""}):
        try:
            info = Path(path).stat()
            versions.append((str(Path(path).resolve()), info.st_size, info.st_mtime_ns))
        except OSError:
            versions.append((path, None, None))
    data = {k: event.get(k) for k in ("pipeline", "recipients", "attachments", "options", "process_overrides")}
    data["source_versions"] = versions
    return hashlib.sha256(json.dumps(data, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()


def prepared_files_for_unit(event, unit):
    files = []
    for recipient in unit.get("recipients") or []:
        name = recipient.get("name", "")
        values = (event.get("process_overrides") or {}).get(name)
        if values is None:
            values = (event.get("process_results") or {}).get(name, [])
        files.extend(split_paths(values) if isinstance(values, str) else values)
    return list(dict.fromkeys(str(path) for path in files))


def prepared_files_current(event):
    if not event.get("process_results") or event.get("process_signature") != processing_signature(event):
        return False
    if any(r.get("name", "") not in (event.get("process_results") or {})
           and r.get("name", "") not in (event.get("process_overrides") or {})
           for r in event.get("recipients") or []):
        return False
    return all(Path(p).is_file() for _, unit in draft_units(event)
               for p in prepared_files_for_unit(event, unit))


def preview_plan(event):
    """产出清单预览（不执行）。返回 [(草稿单元, 说明, [文件名...])]。

    最后一步是「每人一份」时，合并组里每人各有一份，附件包是并起来的。
    """
    steps = [s for s in (event.get("pipeline") or []) if s.get("enabled", True)]
    out = []
    if not steps and not event.get("process_overrides"):
        return out
    prepared = prepared_files_current(event)
    for key, unit in draft_units(event):
        if prepared:
            out.append((key, unit, [Path(p).name for p in prepared_files_for_unit(event, unit)]))
            continue
        names = []
        for it in steps:
            names += _item_preview_files(event, it, unit)
        for a in event.get("attachments") or []:
            if _attachment_matches(a, unit):
                names.append(Path(a.get("path", "")).name or a.get("path", ""))
        for r in unit.get("recipients") or []:
            for p in r.get("attachments") or []:
                names.append(Path(p).name)
        out.append((key, unit, names))
    return out


def run_event(store, event, only_recips=None, log=None, progress=None,
              create_drafts=True, outdir=None, word=None, reuse_processed=False):
    """执行一个发送事件，返回执行记录（Run）。

    只建草稿，绝不发送。执行记录写进 lp_data/runs/。
    word: 可传入外部 WordBatch（GUI 的常驻工作线程用）；不给就用本线程的共享实例。
    """
    started = _now_str()
    signature = processing_signature(event)
    run = {
        "id": _new_id("run"),
        "event_id": event.get("id", ""),
        "event_name": event.get("name", ""),
        "started_at": started,
        "finished_at": "",
        "status": "",
        "items": [],
        "stats": {"ok": 0, "fail": 0, "skip": 0},
        "errors": [],
        "warnings": [],
        "logs": [],
        "only_recips": sorted(only_recips) if only_recips is not None else None,
    }

    def _log(msg):
        run["logs"].append(str(msg))
        if log:
            log(msg)

    errors, warnings, info = preflight(event, store)
    run["errors"] = errors
    run["warnings"] = warnings
    if errors:
        run["status"] = "preflight_failed"
        run["finished_at"] = _now_str()
        _log("预检未通过，已中止（不生成半成品）：%d 个问题" % len(errors))
        ok, save_err = store.save_run(run)
        if not ok:
            run["warnings"].append("执行记录保存失败：%s" % save_err)
        return run

    opts = event.get("options") or {}
    base_out = Path(outdir or opts.get("outdir")
                    or store.get_settings().get("outdir") or app_dir())
    base_out = base_out / sanitize_filename(event.get("name", "事件"), 60)
    base_out.mkdir(parents=True, exist_ok=True)

    # Word 实例按线程复用：每次新建都会残留一个 WINWORD.EXE（见 default_word_batch 注释）
    word = word or default_word_batch()
    cached = bool(reuse_processed and prepared_files_current(event))
    automatic = dict(event, recipients=[r for r in event.get("recipients") or []
                     if r.get("name", "") not in (event.get("process_overrides") or {})])
    if cached or not automatic["recipients"]:
        outputs, pw, perr = {}, [], []
        if cached:
            _log("使用已处理并核对的文件，不重复转换或覆盖自定义对应。")
    else:
        outputs, pw, perr = run_items(automatic, base_out, word=word,
                                     only_recips=only_recips, log=_log, progress=progress)
    run["warnings"] = list(warnings) + list(pw)
    for ukey, sid, msg in perr:
        run["errors"].append("[%s] %s" % (sid, msg))

    # 收集产物：所有条目的产物都进附件包（每个条目内部已把中间产物处理掉，
# 返回的就是最终产物——docx 已转 pdf、pdf 已加水印/加密）
    steps = [s for s in (event.get("pipeline") or []) if s.get("enabled", True)]

    def _outputs_of(item, unit):
        """某条目在某个收件人/合并组下的产物文件。

        条目粒度与发送模式粒度可能不同：
          条目 per_lp + 发送按组合并 → 组内每人各一份，附件包要并起来
          条目 per_group + 发送每人一份 → 取该收件人所在组的那份
        """
        iid = item.get("id")
        cur = outputs.get(iid) or {}
        files = []
        for one in _item_units_for(item, unit):
            files += list(cur.get(_item_output_key(item, one)) or [])
        return files

    def _final_for(unit):
        files = []
        for it in steps:
            files += _outputs_of(it, unit)
        return files

    mail = event.get("mail") or {}
    base_subject = mail.get("subject") or DEFAULT_SUBJECT_V3
    base_body = mail.get("body") or DEFAULT_BODY_V3
    overrides = event.get("mail_overrides") or {}

    for key, unit in draft_units(event):
        _ov = overrides.get(key) or {}
        subject_tpl = _ov.get("subject") or base_subject
        body_tpl = _ov.get("body") or base_body
        names = [r.get("name", "") for r in unit.get("recipients") or []]
        if only_recips is not None and not (set(names) & set(only_recips)):
            continue
        rep = unit.get("rep") or {}
        final_files = _final_for(unit)
        dup_log = []
        if cached:
            files, missing = assemble_attachments(event, unit, prepared_files_for_unit(event, unit),
                                                  dup_log=dup_log, complete=True)
        elif any(r.get("name", "") in (event.get("process_overrides") or {}) for r in unit.get("recipients") or []):
            selected = []
            for recipient in unit.get("recipients") or []:
                name = recipient.get("name", "")
                if name in (event.get("process_overrides") or {}):
                    values = event["process_overrides"][name]
                    selected.extend(split_paths(values) if isinstance(values, str) else values)
                else:
                    one = dict(unit, recipients=[recipient], rep=recipient, lp_name=name,
                               group_key=group_key_of(recipient))
                    prepared, absent = assemble_attachments(event, one, _final_for(one))
                    selected.extend([str(p) for p in prepared] + absent)
            files, missing = assemble_attachments(event, unit, selected, dup_log=dup_log, complete=True)
        else:
            files, missing = assemble_attachments(event, unit, final_files, dup_log=dup_log)
        for first, dup in dup_log:
            _log("[附件重复提示] %s 与 %s 内容相同（不同路径），已全部保留，请人工确认"
                 % (Path(first).name, Path(dup).name))
        item = {
            "key": key,
            "kind": "group" if unit.get("group") else "lp",
            "name": (unit.get("group") or names[0] if names else key),
            "members": names,
            "group": unit.get("group", ""),
            "files": [str(p) for p in files],
            "missing": missing,
            "draft": "",
            "error": "",
            "open_pw_fp": pw_fingerprint(rep.get("open_pw", "")),
        }
        # 本单元的流水线错误
        unit_errors = [msg for ukey, sid, msg in perr
                       if ukey in (key,) or ukey in names]
        if unit_errors:
            item["error"] = "；".join(unit_errors[:3])
        if missing:
            item["draft"] = "跳过(附件缺失)"
            run["stats"]["skip"] += 1
            _log("[附件缺失] %s：%s" % (item["name"], "、".join(missing[:3])))
        elif not files:
            item["draft"] = "跳过(无附件)"
            run["stats"]["skip"] += 1
        elif not unit.get("rep", {}).get("to") and len(names) == 1:
            item["draft"] = "跳过(无邮箱)"
            run["stats"]["skip"] += 1
        else:
            if not create_drafts:
                item["draft"] = "仅生成文件"
                run["stats"]["ok"] += 1
            else:
                try:
                    fmap = build_field_map(rep, event)
                    rec = dict(rep)
                    rec["to"] = ";".join(_merge_addresses(unit["recipients"], "to"))
                    rec["cc"] = ";".join(_merge_addresses(unit["recipients"], "cc",
                                                          mail.get("cc_default", "")))
                    rec["bcc"] = ";".join(_merge_addresses(unit["recipients"], "bcc",
                                                           mail.get("bcc_default", "")))
                    owner_pw = rep.get("owner_pw", "")
                    if opts.get("gp_owner_mode") and opts.get("gp_owner_pw"):
                        owner_pw = opts["gp_owner_pw"]
                    create_outlook_draft(rec, files, subject_tpl, body_tpl,
                                         fmap, owner_pw)
                    item["draft"] = "ok"
                    run["stats"]["ok"] += 1
                    _log("[草稿OK] %s -> Outlook 草稿箱（未发送）" % item["name"])
                except Exception as e:
                    item["draft"] = "失败:%s" % str(e)[:60]
                    item["error"] = str(e)
                    run["stats"]["fail"] += 1
                    _log("[草稿失败] %s：%s" % (item["name"], e))
        if item["error"] and not item["draft"].startswith("失败"):
            run["stats"]["fail"] += 1
        run["items"].append(item)

    if run["stats"]["fail"]:
        run["status"] = "partial"
    elif run["stats"]["skip"]:
        run["status"] = "ok_with_skips"
    else:
        run["status"] = "ok"
    run["finished_at"] = _now_str()
    run["process_signature"] = signature
    run["info"] = info
    ok, save_err = store.save_run(run)
    if not ok:
        # 执行记录写不进去要让人知道（否则「执行历史」里什么都没有却查不到原因）
        run["warnings"].append("执行记录保存失败：%s" % save_err)
        _log("执行记录保存失败：%s" % save_err)
    _log("===== 完成：成功 %d，失败 %d，跳过 %d ====="
         % (run["stats"]["ok"], run["stats"]["fail"], run["stats"]["skip"]))
    return run


def failed_recipients(run):
    """从执行记录里取出需要重跑的收件人名。"""
    names = set()
    for item in (run or {}).get("items") or []:
        status = item.get("draft") or ""
        if status.startswith("失败") or item.get("error"):
            names.update(item.get("members") or [])
    return names


# ============================================================================
# 流程图 / 导入校验 / 导出备份
# ============================================================================


def item_label(item):
    """条目的中文显示名（名称 + 类型 + 关键来源）。"""
    kind = (item or {}).get("kind") or "generate"
    name = (item or {}).get("name") or ITEM_KIND_LABELS.get(kind, "条目")
    params = (item or {}).get("params") or {}
    if kind == "generate":
        tp = params.get("template_path")
        if tp:
            name += "（%s）" % Path(str(tp)).name
    else:
        fl = params.get("files") or []
        if fl:
            name += "（%d 个文件）" % len(fl)
    return name


def flow_nodes(event):
    """把事件配置解析成流程图节点序列（纯逻辑，GUI 只负责画）。

    返回 [{kind, key, title, sub, enabled, granularity}]：
      kind = source(收件人) / step(流水线步骤) / output(最终产物) / draft(Outlook 草稿)
    """
    ev = event or {}
    recips = ev.get("recipients") or []
    mode = send_mode_of(ev)
    nodes = []
    nodes.append({
        "kind": "source", "key": "recipients",
        "title": "收件人",
        "sub": "%d 位 · %s" % (len(recips), SEND_MODE_LABELS.get(mode, mode)),
        "enabled": True, "granularity": mode,
    })
    steps = ev.get("pipeline") or []
    enabled_steps = [s for s in steps if s.get("enabled", True)]
    for i, st in enumerate(steps):
        params = st.get("params") or {}
        gran = st.get("granularity") or "per_lp"
        procs = []
        if params.get("to_pdf"):
            procs.append("PDF")
        if params.get("watermark"):
            procs.append("水印")
        if params.get("encrypt"):
            procs.append("加密")
        sub = " · ".join(procs) if procs else "原样"
        sub += " · " + GRANULARITY_LABELS.get(gran, gran)
        nodes.append({
            "kind": "step",
            "key": st.get("id") or ("item-%d" % i),
            "title": st.get("name") or ITEM_KIND_LABELS.get(st.get("kind"), "条目"),
            "sub": sub,
            "enabled": bool(st.get("enabled", True)),
            "granularity": gran,
            "index": i,
        })
    nodes.append({
        "kind": "output", "key": "__output__",
        "title": "附件包",
        "sub": ("%d 个条目产出" % len(enabled_steps)) if enabled_steps else "无条目",
        "enabled": bool(enabled_steps),
        "granularity": "",
    })
    try:
        nunits = len(draft_units(ev))
    except Exception:
        nunits = 0
    nodes.append({
        "kind": "draft", "key": "__draft__",
        "title": "Outlook 草稿",
        "sub": "%d 封 · 只存草稿不发送" % nunits,
        "enabled": nunits > 0,
        "granularity": "",
    })
    return nodes


# ---------- 导入校验 ----------

_EMAIL_RE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


def _split_addr_list(text):
    """邮箱列表拆分：分号 / 逗号 / 空格 / 换行都算分隔。"""
    parts = re.split(r"[;,\s]+", str(text or "").strip())
    return [p for p in parts if p]


def validate_recipient_rows(rows, existing_names=None, check_files=True):
    """校验导入的收件人数据。

    返回 {
      "ok": bool,                # 有没有致命错误（缺表头列等）
      "fatal": [str],           # 整表级别的错误
      "items": [{               # 每行一条
         "no": 行号(从 2 开始，1 是表头),
         "name": 名称,
         "status": "ok" / "error" / "warning",
         "errors": [str],
         "warnings": [str],
         "record": 解析出的收件人 dict,
      }],
      "error_count": int, "warning_count": int,
      "unknown_cols": [str],    # 未识别列（会变成自定义字段）
      "field_cols": [str],
      "missing_cols": [str],    # 缺失的关键列
      "duplicate_names": [str],
    }
    """
    res = {
        "ok": True, "fatal": [], "items": [], "notes": [],
        "error_count": 0, "warning_count": 0,
        "unknown_cols": [], "field_cols": [],
        "missing_cols": [], "duplicate_names": [],
    }
    if not rows:
        res["ok"] = False
        res["fatal"].append("文件是空的（读不到任何行）")
        return res

    header = ["" if h is None else str(h).strip() for h in rows[0]]
    lower = [h.lower() for h in header]
    known = set()
    for aliases in HEADER_ALIASES.values():
        known |= {a.lower() for a in aliases}
    for i, h in enumerate(lower):
        if h and h not in known:
            res["unknown_cols"].append(header[i])
    if not any(h in known for h in lower if h):
        res["ok"] = False
        res["fatal"].append(
            "表头里没有识别到任何已知列。至少要有「收件人名称」列；"
            "可识别的列名：%s" % "、".join(sorted(known)))
        return res
    if not any(h in HEADER_ALIASES["name"] for h in lower if h):
        res["ok"] = False
        res["missing_cols"].append("收件人名称")
        res["fatal"].append("缺少「收件人名称」列，无法识别每一行是哪位收件人")
        return res

    recips, _unknown, field_cols = parse_import_rows_v3(rows)
    res["unknown_cols"] = _unknown
    res["field_cols"] = field_cols

    # 空名行在 parse 阶段被丢弃（不导入——这是既定逻辑，不改）。
    # 这里把它们补回结果表：显示为红色 error（名称为空，该行不会导入），
    # 让用户在导入结果里看得见，而不是静默消失。完全空白的行不显示（无意义）。
    # 同时修正行号：丢行后 enumerate 会错位，改为按原始行号对齐。
    name_idx = None
    if rows:
        for idx, h in enumerate(rows[0]):
            hs = "" if h is None else str(h).strip().lower()
            if hs in HEADER_ALIASES.get("name", ()):
                name_idx = idx
                break
    kept_nos, empty_items = [], []
    for ridx, row in enumerate(rows[1:] if rows else []):
        def _cell(i):
            if i is None or i >= len(row):
                return ""
            v = row[i]
            return "" if v is None else str(v).strip()
        nm = _cell(name_idx)
        if nm and not nm.startswith("示例："):
            kept_nos.append(ridx + 2)
        elif not nm and any((str(c).strip() if c is not None else "")
                            for c in row):
            empty_items.append({
                "no": ridx + 2, "name": "", "status": "error",
                "errors": ["名称为空（该行不会导入）"], "warnings": [],
                "record": {}})
    existing = set(existing_names or [])
    seen = {}
    has_owner_pw = False
    for idx, rec in enumerate(recips):
        no = kept_nos[idx] if idx < len(kept_nos) else idx + 2
        errors, warnings = [], []
        name = rec.get("name", "")
        if not name:
            errors.append("名称为空")
        if name in seen:
            errors.append("与本表第 %d 行重名" % seen[name])
            res["duplicate_names"].append(name)
        else:
            seen[name] = no
        if name and name in existing:
            warnings.append("与已有收件人同名，导入后会覆盖更新")

        to = rec.get("to", "")
        if not to:
            errors.append("收件人邮箱为空")
        else:
            for addr in _split_addr_list(to):
                if not _EMAIL_RE.match(addr):
                    errors.append("收件人邮箱格式不对：%s" % addr)
        for key, label in (("cc", "抄送"), ("bcc", "密送")):
            val = rec.get(key, "")
            if val:
                for addr in _split_addr_list(val):
                    if not _EMAIL_RE.match(addr):
                        errors.append("%s邮箱格式不对：%s" % (label, addr))

        ow = rec.get("open_pw", "")
        op = rec.get("owner_pw", "")
        if ow and op and ow == op:
            errors.append("打开口令与权限口令相同（违反铁律，必须是两个不同的口令）")
        if op:
            has_owner_pw = True
        if ow and len(ow) < 6:
            warnings.append("打开口令少于 6 位，建议改长")

        files = rec.get("attachments") or []
        if check_files:
            for p in files:
                if not Path(str(p)).exists():
                    warnings.append("附件文件不存在：%s" % p)
        if rec.get("group"):
            warnings.append("已指定合并组「%s」：只有发送模式为「按组合并」时才生效"
                            % rec.get("group"))

        if errors:
            status = "error"
        elif warnings:
            status = "warning"
        else:
            status = "ok"
        res["error_count"] += len(errors)
        res["warning_count"] += len(warnings)
        res["items"].append({
            "no": no, "name": name, "status": status,
            "errors": errors, "warnings": warnings, "record": rec,
        })

    # 空名行并入结果（红色 error），按行号排序；计数在 return 前统一补
    res["items"].extend(empty_items)
    res["items"].sort(key=lambda it: it.get("no", 0))
    res["error_count"] += len(empty_items)

    if has_owner_pw:
        res["notes"].append(
            "文件里填写了权限口令：只用于文件权限控制，不会进入邮件正文或标题。")
    if not res["items"]:
        res["ok"] = False
        res["fatal"].append("没有解析到任何收件人记录（表头下全是空行）")
    return res


# ---------- 导出 ----------

POOL_TAG_HEADERS = {"标签", "联系人标签", "收件人标签", "tag", "tags"}


def parse_pool_tags(value):
    """库标签：支持常用分隔符和导出时对特殊标签使用的 JSON 列表。"""
    if isinstance(value, list):
        values = value
    else:
        text = str(value or "").strip()
        try:
            values = json.loads(text) if text.startswith("[") else None
        except (ValueError, TypeError):
            values = None
        if not isinstance(values, list):
            values = re.split(r"[;,；，、\r\n]+", text)
    tags = []
    for value in values:
        tag = str(value).strip()
        if tag and tag not in tags:
            tags.append(tag)
    return tags


def validate_pool_recipient_rows(rows, existing_names=None, check_files=True):
    """库专用导入：识别标签，明确忽略事件字段，不改变事件导入语义。"""
    result = validate_recipient_rows(rows, existing_names, check_files)
    tag_cols = [c for c in result["field_cols"] if c.casefold() in POOL_TAG_HEADERS]
    ignored = [c for c in result["field_cols"] if c not in tag_cols]
    result["unknown_cols"] = []
    result["field_cols"] = []
    result["ignored_cols"] = ignored
    if ignored:
        result["notes"].append("以下列属于事件数据，本次不导入收件人库：" + "、".join(ignored))
    result["notes"].append("同名联系人只补空缺，已有口令保留；标签取并集，空白标签不会清空原标签。")
    for item in result["items"]:
        rec = item.get("record") or {}
        fields = rec.get("fields") or {}
        rec["tags"] = parse_pool_tags([tag for col in tag_cols
                                       for tag in parse_pool_tags(fields.get(col, ""))])
        rec["fields"] = {}
        for i, warning in enumerate(item["warnings"]):
            if "与已有收件人同名" in warning:
                item["warnings"][i] = "与已有联系人同名，按名称补空合并；已有口令保留，标签合并"
    return result

EXPORT_BASE_HEADERS = ["收件人名称", "水印名称", "收件人", "抄送", "密送",
                       "邮件标题", "附件路径", "合并组", "打开口令", "权限口令"]


def import_template_headers(extra=()):
    """「收件人清单导入模板」的表头 —— 不含「附件路径」。

    为什么要和导出表头分开：导出是数据备份（旧数据的 attachments 路径要留住），
    而导入模板是发给业务方填的，v3.4.0 起收件人页已不再露出「附件路径」，
    模板里也不该再出现；v3.4.4 起导入遇到该列也只收进自定义字段。
    """
    base = [h for h in EXPORT_BASE_HEADERS if h != "附件路径"]
    return base + list(extra)


def recipients_to_rows(recipients, include_passwords=True):
    """收件人列表 → 二维列表（可直接写成 Excel）。

    自定义字段列并到末尾，列名取所有收件人的字段键的并集（按首次出现顺序）。
    权限口令默认导出（这是备份用途），调用方可以传 include_passwords=False 抹掉。
    """
    recips = recipients or []
    extra = []
    for r in recips:
        for k in (r.get("fields") or {}):
            if k not in extra:
                extra.append(k)
    headers = list(EXPORT_BASE_HEADERS) + extra
    rows = [headers]
    for r in recips:
        # attachments 是**路径字符串列表**（不是 dict 列表）——这里必须按字符串取，
        # 写成 a.get("path") 会 AttributeError，导出按钮会静默失败。
        att_cells = []
        for a in (r.get("attachments") or []):
            if isinstance(a, str):
                att_cells.append(a)
            elif isinstance(a, dict):
                att_cells.append(str(a.get("path", "")))
        row = [
            r.get("name", ""), r.get("wm", ""), r.get("to", ""),
            r.get("cc", ""), r.get("bcc", ""), r.get("subject", ""),
            ";".join(x for x in att_cells if x),
            r.get("group", ""),
            r.get("open_pw", "") if include_passwords else "",
            r.get("owner_pw", "") if include_passwords else "",
        ]
        fields = r.get("fields") or {}
        for k in extra:
            row.append(fields.get(k, ""))
        rows.append(row)
    return rows


def pool_recipients_to_rows(recipients, include_passwords=True):
    """库 Excel：基础资料加标签，排除历史事件字段；标签可无损往返。"""
    recs = [dict(r, fields={}) for r in recipients]
    rows = recipients_to_rows(recs, include_passwords)
    rows[0].append("标签")
    for row, rec in zip(rows[1:], recs):
        tags = parse_pool_tags(rec.get("tags") or [])
        special = any(re.search(r"[;,；，、\r\n]", t) or t.startswith("[") for t in tags)
        row.append(json.dumps(tags, ensure_ascii=False) if special else "；".join(tags))
    return rows


def events_recipients_to_rows(events, include_passwords=True):
    """跨事件明细：所有事件共用字段并集表头，各行按同一列顺序填值。"""
    names, recs = [], []
    for event in events:
        for rec in event.get("recipients") or []:
            names.append(event.get("name", ""))
            recs.append(rec)
    headers = ["所属事件"] + list(EXPORT_BASE_HEADERS)
    keys = list(dict.fromkeys(k for rec in recs for k in (rec.get("fields") or {})))
    for key in keys:
        label = key if key not in headers else "本事件字段：" + key
        original, suffix = label, 2
        while label in headers:
            label = "%s（%d）" % (original, suffix)
            suffix += 1
        headers.append(label)
    rows = [headers]
    for name, rec in zip(names, recs):
        base = recipients_to_rows([dict(rec, fields={})], include_passwords)[1]
        fields = rec.get("fields") or {}
        rows.append([name] + base + [fields.get(key, "") for key in keys])
    return rows


def event_export_payload(event):
    """事件导出（备份用）。"""
    return {
        "format": "doc-mailer.event",
        "format_version": 1,
        "exported_at": _now_str(),
        "app_version": APP_VERSION,
        "event": event or {},
    }


def event_import_payload(payload):
    """解析事件备份文件。返回 (ok, event, error)。"""
    if not isinstance(payload, dict):
        return False, None, "文件内容不是有效的 JSON 对象"
    ev = payload.get("event")
    if not isinstance(ev, dict):
        return False, None, "文件里没有 event 字段（不是本程序导出的事件备份）"
    if not ev.get("name"):
        return False, None, "事件缺少名称"
    return True, normalize_event(ev), ""


def backup_payload(store, include_runs=True):
    """整库备份内容（不含 runs 时只备设置/事件/场景）。"""
    data = {
        "format": "doc-mailer.backup",
        "format_version": 1,
        "exported_at": _now_str(),
        "app_version": APP_VERSION,
        "settings": store.load_settings()[0] or {},
        "events": [],
        "scenarios": [],
        "runs": [],
    }
    for ev in store.list_events():
        data["events"].append(ev)
    for sc in store.list_scenarios():
        data["scenarios"].append(sc)
    if include_runs:
        for r in store.list_runs():
            data["runs"].append(r)
    return data

