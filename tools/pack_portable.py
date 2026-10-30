# -*- coding: utf-8 -*-
"""打便携 zip（顶层 app/），并读回 zip 内 lp_tool.py 断言版本与关键函数。

用法：python pack_portable.py [版本]

与 HANDOFF §7 步骤 3 的差异（必要修正，不是自创流程）：
  §7 的示例用 os.walk 把绿色版目录整个打进去，但绿色版里常驻着
  lp_data/（用户自己的测试数据）与 __pycache__/，两者都**不该进包**。
  这里显式排除它们，其余结构与 §7 一致（顶层 app/）。
"""
import os
import re
import sys
import zipfile
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if ROOT.name == 'tools' and ROOT.parent.name == 'lp-doc-mailer':
    ROOT = ROOT.parent.parent
PKG = ROOT / "LP_Tool_绿色版"


def _app_version():
    """版本号从 lp_tool.py 里读，别写死 —— 写死过一次就漏发过一次。"""
    src = (ROOT / "lp_tool.py").read_text(encoding="utf-8")
    m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', src)
    if not m:
        print("[X] lp_tool.py 里找不到 APP_VERSION")
        sys.exit(1)
    return m.group(1)


VERSION = sys.argv[1] if len(sys.argv) > 1 else _app_version()
OUT = ROOT / ("doc-mailer_v%s_portable.zip" % VERSION)
UPDATE_OUT = ROOT / ("doc-mailer_v%s_update.zip" % VERSION)
EXCLUDE_DIRS = {"lp_data", "__pycache__", ".pytest_cache"}
EXCLUDE_SUFFIX = (".pyc", ".pyo")
RUNTIME_MARKER = "app/runtime/.docmailer-runtime-id"
UPDATE_FILES = ("lp_tool.py", "lp_v3.py", "watermark_encrypt.py", "start.bat", "icon.ico", "使用说明.txt", "源码获取说明.txt")


def runtime_fingerprint():
    root = PKG / "runtime"
    digest = hashlib.sha256()
    files = sorted(p for p in root.rglob('*') if p.is_file() and
                   not any(part in EXCLUDE_DIRS for part in p.relative_to(root).parts) and
                   not p.name.endswith(EXCLUDE_SUFFIX) and p.name != '.docmailer-runtime-id')
    for path in files:
        relative = path.relative_to(root).as_posix().encode('utf-8')
        digest.update(len(relative).to_bytes(4, 'big')); digest.update(relative)
        file_hash = hashlib.sha256()
        with path.open('rb') as stream:
            for block in iter(lambda: stream.read(1048576), b''):
                file_hash.update(block)
        digest.update(file_hash.digest())
    return digest.hexdigest()

# zip 内必须存在的关键内容（用来证明打进去的是对的东西）
MUST_HAVE = [
    "app/lp_tool.py",
    "app/lp_v3.py",
    "app/watermark_encrypt.py",
    "app/start.bat",
    "app/lp_data.example.json",
    "app/runtime/pythonw.exe",
    "app/使用说明.txt",
]
MUST_NOT_HAVE = ["app/lp_data/", "__pycache__"]


def build():
    if not PKG.is_dir():
        print("[X] 绿色版目录不存在：%s" % PKG)
        return 1
    n = 0
    runtime_id = runtime_fingerprint()
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for root, dirs, files in os.walk(PKG):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            rel_root = os.path.relpath(root, PKG)
            if rel_root == ".":
                rel_root = ""
            for f in files:
                if f.endswith(EXCLUDE_SUFFIX) or (rel_root.replace('\\', '/') == 'runtime' and f == '.docmailer-runtime-id'):
                    continue
                fp = os.path.join(root, f)
                arc = os.path.join("app", rel_root, f) if rel_root else os.path.join("app", f)
                arc = arc.replace("\\", "/")
                z.write(fp, arc)
                n += 1
        z.writestr(RUNTIME_MARKER, runtime_id + '\n')
    with zipfile.ZipFile(UPDATE_OUT, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for name in UPDATE_FILES:
            if (PKG / name).is_file():
                z.write(PKG / name, 'app/' + name)
        z.writestr('app/.update_manifest.json', json.dumps({'version': VERSION, 'runtime_id': runtime_id}))
    print('[OK] 轻量更新包：%s（%.2f MB）' % (UPDATE_OUT.name, UPDATE_OUT.stat().st_size / 1048576))
    print("[OK] 打包完成：%s（%d 个文件，%.1f MB）"
          % (OUT.name, n, OUT.stat().st_size / 1024 / 1024))
    return 0


def verify():
    print("--- 读 zip 内断言 ---")
    with zipfile.ZipFile(OUT) as z:
        names = z.namelist()
        bad = z.testzip()
        if bad:
            print("[X] zip 损坏，首个坏文件：%s" % bad)
            return 1
        for need in MUST_HAVE:
            if need not in names:
                print("[X] 缺少必需文件：%s" % need)
                return 1
        for junk in MUST_NOT_HAVE:
            hit = [x for x in names if x.startswith(junk)]
            if hit:
                print("[X] 混进了不该有的内容：%s（%d 项）" % (junk, len(hit)))
                return 1
        src = z.read("app/lp_tool.py").decode("utf-8")
        assert 'APP_VERSION = "%s"' % VERSION in src, "zip 内 APP_VERSION 不是 %s" % VERSION
        print("[OK] app/lp_tool.py 里 APP_VERSION = \"%s\"" % VERSION)
        v3 = z.read("app/lp_v3.py").decode("utf-8")
        assert 'APP_VERSION = "%s"' % VERSION in v3, "zip 内 lp_v3.py 版本不一致"
        for fn in ("def normalize_item", "def execute_item", "def run_items",
                   "def field_source_options", "def default_word_batch",
                   "def normalize_pool_recipient", "def pool_merge_recipient",
                   "def merge_recipients_into_event"):
            assert fn in v3, "lp_v3.py 缺少关键函数 %s" % fn
        print("[OK] lp_v3.py 关键函数齐全（normalize_item/execute_item/run_items/"
              "field_source_options/default_word_batch/收件人库三件套）")
        for fn in ("def _wb_step_pipeline", "def _wb_step_overview",
                   "def _wb_item_params", "def _wb_fill_pipe_table",
                   "def _refresh_pool_tab", "def _refresh_runs_tab",
                   "def _recipient_dialog"):
            assert fn in src, "lp_tool.py 缺少关键函数 %s" % fn
        print("[OK] lp_tool.py 关键函数齐全（流水线表格/总览页/参数区/收件人库页/执行历史页）")
        guide = z.read("app/使用说明.txt").decode("utf-8")
        assert "文档分发助手 v%s" % VERSION in guide, \
            "zip 内 使用说明.txt 的版本号不是 v%s" % VERSION
        assert "文档流水线" in guide and "总览" in guide, \
            "zip 内 使用说明.txt 不是 v3 版"
        assert "收件人库" in guide and "执行历史" in guide, \
            "zip 内 使用说明.txt 缺 P4 的收件人库/执行历史说明"
        assert "①页" not in guide and "②页" not in guide, \
            "zip 内 使用说明.txt 还残留 v2 的「①页/②页」旧说法"
        print("[OK] 使用说明.txt 是 v%s 版（六步工作台/流水线/总览/收件人库/执行历史），无 v2 残留" % VERSION)
        print("[OK] 文件总数 %d，无 lp_data/、无 __pycache__/" % len(names))
        marker = z.read(RUNTIME_MARKER).decode('ascii').strip()
    with zipfile.ZipFile(UPDATE_OUT) as z:
        assert z.testzip() is None
        manifest = json.loads(z.read('app/.update_manifest.json'))
        assert manifest == {'version': VERSION, 'runtime_id': marker}
        assert not any(n.startswith('app/runtime/') or n.startswith('app/lp_data/') for n in z.namelist())
        assert all('app/' + n in z.namelist() for n in UPDATE_FILES[:4])
        print('[OK] 轻量包版本及运行库标识匹配，未包含运行库与用户数据')
    return 0


if __name__ == "__main__":
    rc = build()
    if rc:
        sys.exit(rc)
    sys.exit(verify())
