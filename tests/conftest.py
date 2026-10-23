# -*- coding: utf-8 -*-
"""pytest 公共夹具。

Tk 根窗口在整个测试进程里只创建一次：Windows 上「创建 → 销毁 → 再创建」会偶发
Tcl 库重初始化失败（tk.tcl 读不到 ttk/sizegrip.tcl）。应用本身每进程只建一个 root，
不受影响；这里统一共享根窗口即可。
"""
import pytest

tk = pytest.importorskip("tkinter")


@pytest.fixture(scope="session")
def tk_root():
    root = tk.Tk()
    root.withdraw()
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
    """
    pytest.importorskip("win32com.client")
    import lp_v3
    return lp_v3.default_word_batch()


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
