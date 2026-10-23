# -*- coding: utf-8 -*-
"""pytest 公共夹具。

Tk 根窗口在整个测试进程里只创建一次：Windows 上「创建 → 销毁 → 再创建」会偶发
Tcl 库重初始化失败（tk.tcl 读不到 ttk/sizegrip.tcl）。应用本身每进程只建一个 root，
不受影响；这里统一共享根窗口即可。

**为什么夹具要能「优雅跳过」**：
CI（GitHub Actions windows-latest）上没有预装 Microsoft Office，也没有交互式
桌面会话。Word 相关测试（生成 docx / 转 PDF）与 GUI 测试一旦在那种环境里硬跑，
会直接 error 而不是 skip —— error 会让整个 CI workflow 失败，Release 就发不出来。
所以这里做成「环境不具备就 skip」，让 CI 上照样绿、只是少跑一部分用例。

本机想验证 CI 路径时，可以设环境变量强制跳过：
    $env:LP_SKIP_WORD_TESTS='1'    # 模拟「没装 Word」
"""
import os

import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture(scope="session")
def tk_root():
    """共享的 Tk 根窗口；没有可用显示环境时整组 GUI 测试跳过。"""
    try:
        root = tk.Tk()
        root.withdraw()
    except Exception as e:      # 无显示环境 / Tcl 初始化失败
        pytest.skip("无法创建 Tk 根窗口（无显示环境？）：%s" % e)
    yield root
    try:
        root.destroy()
    except Exception:
        pass


@pytest.fixture(scope="session")
def word_batch():
    """整个测试进程共用一个 Word 实例。

    复用 lp_v3.default_word_batch()：每次新建 WordBatch 都会 DispatchEx 出一个
    独立的 WINWORD.EXE（close() 只释放引用、不 Quit），反复新建会把机器拖垮
    （实测堆积过 71 个）。进程级只保留一个即可。

    本机没装 Word 时跳过（CI runner 就是这种情况），否则用例会 error 而不是 skip。
    """
    pytest.importorskip("win32com.client")
    if os.environ.get("LP_SKIP_WORD_TESTS"):
        pytest.skip("LP_SKIP_WORD_TESTS 已设置：跳过需要 Word 的用例")
    import lp_v3
    wb = lp_v3.default_word_batch()
    try:
        wb._ensure()      # 探测 Word COM 是否真的可用（WordBatch 是懒加载的）
    except Exception as e:
        pytest.skip("本机没有可用的 Microsoft Word：%s" % e)
    return wb


@pytest.fixture(scope="module")
def word_template(tmp_path_factory, word_batch):
    """用 Word 造一个带占位符的 .docx 模板（需要本机装了 Word）。"""
    d = tmp_path_factory.mktemp("word")
    path = d / "tpl.docx"
    word_batch.create_docx(path,
                            "Capital Call Notice\r\r"
                            "Dear {{LP_Name}},\r\r"
                            "Call Amount: USD {{Call_Amount}}\r"
                            "Due Date: {{Due_Date}}\r")
    return path
