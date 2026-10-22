# -*- coding: utf-8 -*-
"""核心纯逻辑回归测试（无需 GUI/Outlook）：pytest tests/"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz
import pytest

from lp_tool import (MOCK_HEADERS, MOCK_LPS, DEFAULT_FILENAME_TPL, gen_pw,
                     parse_import_rows, render_filename, render_template,
                     sanitize_filename, split_paths, split_mails,
                     validate_records, pw_fingerprint)
from watermark_encrypt import process_one


def test_parse_mock_rows():
    lps, unknown = parse_import_rows([MOCK_HEADERS] + MOCK_LPS)
    assert len(lps) == 8 and unknown == []
    assert all(r["open_pw"] == "" or r["open_pw"] for r in lps)  # 空值保留语义
    # 示例占位行跳过
    lps2, _ = parse_import_rows([MOCK_HEADERS, ["示例：公司", "（留空）", "a@x.com"]])
    assert lps2 == []


def test_parse_keeps_empty_passwords():
    rows = [MOCK_HEADERS, ["甲", "", "a@x.com", "", "", "", "", "", ""]]
    lps, _ = parse_import_rows(rows)
    assert lps[0]["open_pw"] == "" and lps[0]["owner_pw"] == ""


def test_validate_records():
    bad = [{"name": "甲", "wm": "x", "open_pw": "1", "owner_pw": "1"},
           {"name": "乙", "wm": "x", "open_pw": "1", "owner_pw": "2"},
           {"name": "丙", "wm": " ", "open_pw": "", "owner_pw": "3"}]
    issues = validate_records(bad)
    assert any("相同" in i for i in issues)
    assert any("重复" in i for i in issues)
    assert any("水印名称为空" in i for i in issues)
    assert validate_records([{"name": "甲", "wm": "x", "open_pw": "a1",
                              "owner_pw": "b2"}]) == []


def test_template_escape():
    assert render_template("{{收件人名称}}", {"name": "A&B<x>"}) == "A&amp;B&lt;x&gt;"
    assert render_template("{{打开口令}}", {"open_pw": "P<w&>"
                                            }, escape=False) == "P<w&>"


def test_split_helpers():
    assert split_paths("C:/a.pdf;" + chr(10) + "D:/b,c.pdf") == ["C:/a.pdf", "D:/b,c.pdf"]
    assert split_mails("a@x.com" + chr(10) + "b@x.com") == ["a@x.com", "b@x.com"]
    assert split_mails("张三 <z@x.com>") == ["张三 <z@x.com>"]


def test_sanitize_and_filename():
    assert sanitize_filename('A<B>C:D*E?F"G|H/I' + chr(10) + "J.") == "A_B_C_D_E_F_G_H_IJ"
    assert render_filename(DEFAULT_FILENAME_TPL, "/d/年报.pdf",
                           {"name": "甲公司"}) == "年报_甲公司"
    assert pw_fingerprint("x") and len(pw_fingerprint("x")) == 8


def test_html_roundtrip_idempotent():
    import tkinter as tk
    from lp_tool import load_html, text_to_html
    root = tk.Tk(); root.withdraw()
    try:
        tw = tk.Text(root)
        tw.insert("1.0", "第一段" + chr(10) * 2 + "第三段")
        prev = text_to_html(tw)
        for _ in range(3):
            load_html(tw, prev)
            prev = text_to_html(tw)
        load_html(tw, prev)
        assert tw.get("1.0", "end-1c") == "第一段" + chr(10) * 2 + "第三段"
    finally:
        root.destroy()


def test_process_one_guards(tmp_path):
    src = tmp_path / "s.pdf"
    doc = fitz.open(); doc.new_page(); doc.save(src); doc.close()
    with pytest.raises(ValueError):
        process_one(src, tmp_path / "o.pdf", "  ", "a123456789", "b123456789", True)
    with pytest.raises(ValueError):
        process_one(src, tmp_path / "o.pdf", "水印", "same", "same", True)
    with pytest.raises(ValueError):
        process_one(src, tmp_path / "o.pdf", "水印", "", "b123456789", True)


def test_process_one_permissions(tmp_path):
    src = tmp_path / "s.pdf"
    doc = fitz.open(); doc.new_page(); doc.save(src); doc.close()
    out = tmp_path / "enc.pdf"
    info = process_one(src, out, "测试水印", "Open123456", "Owner123456", False)
    assert info["encrypted"]
    assert info["user_perms"]["打印"] is True
    assert info["user_perms"]["修改"] is False and info["user_perms"]["复制文字"] is False
    assert not (tmp_path / "enc.pdf.tmp").exists()
