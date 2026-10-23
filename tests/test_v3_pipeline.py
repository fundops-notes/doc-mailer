# -*- coding: utf-8 -*-
"""v3.1 流水线回归测试（两种产出条目 / 字段级来源 / 权限位 / Word 格式保持）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz
import pytest

from lp_v3 import (RunContext, build_field_map, build_units, field_source_options,
                   group_key_of, item_output_name, item_placeholders,
                   item_uses_sheet, load_sheet_data, new_event, new_item,
                   normalize_item, normalize_recipient, preview_plan,
                   render_with_fields, run_items, scrub_owner_password,
                   sheet_to_dicts)
import watermark_encrypt as wm


def _event(name="E", recips=None, pipeline=None, **opts):
    ev = new_event(name)
    ev["recipients"] = [normalize_recipient(r) for r in (recips or [])]
    if pipeline is not None:
        ev["pipeline"] = [normalize_item(s) for s in pipeline]
    ev["options"].update(opts)
    return ev


def _attach(files=None, **params):
    p = {"files": list(files or [])}
    p.update(params)
    return {"kind": "attach", "params": p}


def _generate(template=None, mapping=None, name="", **params):
    p = {"template_path": template or ""}
    p.update(params)
    return {"kind": "generate", "name": name, "params": p, "mapping": mapping or {}}


# ---------------------------------------------------------------- 字段映射

def test_field_map_never_contains_owner_password():
    """铁律：权限口令绝不进字段映射。"""
    rec = normalize_recipient({"name": "甲", "open_pw": "op", "owner_pw": "SECRET_OWNER"})
    fmap = build_field_map(rec, new_event("E"))
    assert "{{打开口令}}" in fmap
    assert "SECRET_OWNER" not in "".join(fmap.values())
    for k in fmap:
        assert "owner" not in k.lower()
        assert "权限口令" not in k


def test_field_map_rejects_forbidden_custom_column():
    rec = normalize_recipient({"name": "甲", "fields": {"权限口令": "x", "Call": "1"}})
    fmap = build_field_map(rec, new_event("E"))
    assert "{{Call}}" in fmap
    assert "{{权限口令}}" not in fmap


def test_field_map_sources():
    ev = new_event("E")
    ev["options"]["constants"] = {"Fund_Name": "Linear Capital VI L.P."}
    rec = normalize_recipient({"name": "甲", "fields": {"Call_Amount": "200,000"}})
    fmap = build_field_map(rec, ev)
    assert fmap["{{Call_Amount}}"] == "200,000"
    assert fmap["{{fields.Call_Amount}}"] == "200,000"
    assert fmap["{{Fund_Name}}"] == "Linear Capital VI L.P."
    assert fmap["{{收件人名称}}"] == "甲"
    assert fmap["{{name}}"] == "甲"


def test_render_with_fields_supports_both_placeholder_styles():
    fmap = {"{{LP_Name}}": "甲公司", "{{Call_Amount}}": "200,000"}
    out = render_with_fields("Dear {{LP_Name}}, {{Call_Amount}} / ?LP_Name?", fmap)
    assert out == "Dear 甲公司, 200,000 / 甲公司"


def test_render_scrubs_owner_password_even_if_user_typed_it():
    fmap = {"{{x}}": "SECRET_OWNER"}
    out = render_with_fields("口令是 {{x}}", fmap, owner_pw="SECRET_OWNER")
    assert "SECRET_OWNER" not in out
    assert "已屏蔽" in out


def test_scrub_owner_password_noop_when_empty():
    assert scrub_owner_password("abc", "") == "abc"


# ---------------------------------------------------------------- 条目模型

def test_normalize_item_defaults():
    it = normalize_item({})
    assert it["kind"] == "generate"
    assert it["enabled"] is True
    assert it["granularity"] == "per_lp"
    assert it["params"]["to_pdf"] is False
    assert it["params"]["watermark"] is False
    assert it["params"]["encrypt"] is False
    assert it["params"]["permissions"] == ["print"]
    assert it["id"]


def test_normalize_item_rejects_unknown_kind():
    it = normalize_item({"kind": "pdf_watermark"})
    assert it["kind"] == "generate", "旧的六种步骤类型不再出现"


def test_new_item_gives_default_name():
    assert new_item("generate")["name"] == "新生成文档"
    assert new_item("attach")["name"] == "新附加文件"


def test_item_ids_are_unique():
    a = normalize_item({})
    b = normalize_item({})
    assert a["id"] != b["id"]


def test_normalize_item_keeps_explicitly_empty_permissions():
    """用户主动清空权限位 = 只能打开查看（区别于「没配过」）。"""
    it = normalize_item({"params": {"permissions": []}})
    assert it["params"]["permissions"] == []


def test_normalize_item_reads_legacy_files_param_string():
    it = normalize_item({"kind": "attach", "params": {"files": "a.pdf;b.pdf"}})
    assert it["params"]["files"] == ["a.pdf", "b.pdf"]


def test_item_uses_sheet_only_for_sheet_sources():
    it = normalize_item({"mapping": {"{{X}}": "fields.X"}})
    assert item_uses_sheet(it) is False
    it2 = normalize_item({"mapping": {"{{X}}": "sheet.X"}})
    assert item_uses_sheet(it2) is True


def test_field_source_options_lists_sheet_cols_only_when_given():
    ev = _event(recips=[{"name": "甲", "fields": {"Call": "1"}}])
    rec = ev["recipients"][0]
    plain = field_source_options(ev, rec, None)
    assert not any(k.startswith("sheet.") for k, _c, _l in plain)
    assert ("fields.Call", "收件人字段", "Call") in plain
    with_sheet = field_source_options(ev, rec, ["Due_Date"])
    assert ("sheet.Due_Date", "外部数据表", "Due_Date") in with_sheet


# ---------------------------------------------------------------- 产出单元

def test_build_units_per_lp():
    ev = _event(recips=[{"name": "甲"}, {"name": "乙"}])
    units = build_units(ev, "per_lp")
    assert [u[0] for u in units] == ["甲", "乙"]


def test_build_units_per_group_merges():
    ev = _event(recips=[{"name": "甲", "group": "G1"}, {"name": "乙", "group": "G1"},
                        {"name": "丙"}])
    units = build_units(ev, "per_group")
    keys = [u[0] for u in units]
    assert "G1" in keys
    assert len([k for k in keys if k.startswith("§")]) == 1  # 无组的自成一组
    g1 = dict(units)["G1"]
    assert len(g1["recipients"]) == 2
    assert g1["rep"]["name"] == "甲"


def test_build_units_all_is_single():
    ev = _event(recips=[{"name": "甲"}, {"name": "乙"}])
    units = build_units(ev, "all")
    assert len(units) == 1 and units[0][0] == "__all__"
    assert units[0][1]["rep"] is None


def test_group_key_of_ungrouped():
    assert group_key_of({"name": "甲"}) == "§甲"
    assert group_key_of({"name": "甲", "group": "G"}) == "G"


# ---------------------------------------------------------------- 条目执行

def _make_pdf(path, text="hello"):
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), text)
    doc.save(str(path))
    doc.close()
    return path


def test_attach_item_watermark_then_encrypt(tmp_path):
    """附加文件条目：PDF → 水印 → 加密（一条目内顺序固定）。"""
    src = _make_pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲公司", "wm": "甲公司专属水印",
                         "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                pipeline=[_attach([str(src)], watermark=True, encrypt=True)])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    out = outputs[iid]["甲公司"][0]
    assert out.exists()

    d = fitz.open(str(out))
    try:
        assert d.needs_pass, "输出必须是加密的"
        assert d.authenticate("Open123456") >= 2
        perms = d.permissions
        text = d[0].get_text()
    finally:
        d.close()
    # 权限位：默认仅打印
    assert perms & fitz.PDF_PERM_PRINT
    assert not (perms & fitz.PDF_PERM_MODIFY)
    assert not (perms & fitz.PDF_PERM_COPY)
    assert "甲公司专属水印" in text


def test_encrypt_rejects_same_password(tmp_path):
    src = _make_pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲", "open_pw": "same123456", "owner_pw": "same123456"}],
                pipeline=[_attach([str(src)], encrypt=True)])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert errs and "相同" in errs[0][2]


def test_permissions_are_configurable(tmp_path):
    """v3.1：权限位可放开（默认仍是仅打印）。"""
    src = _make_pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲", "open_pw": "Open123456", "owner_pw": "Owner123456"}],
                pipeline=[_attach([str(src)], encrypt=True,
                                  permissions=["print", "edit", "copy"])])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    out = outputs[iid]["甲"][0]
    d = fitz.open(str(out))
    try:
        assert d.authenticate("Open123456") >= 2
        perms = d.permissions
    finally:
        d.close()
    assert perms & fitz.PDF_PERM_PRINT
    assert perms & fitz.PDF_PERM_MODIFY
    assert perms & fitz.PDF_PERM_COPY


def test_watermark_uses_per_recipient_text(tmp_path):
    src = _make_pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲", "wm": "甲的水印", "open_pw": "o123456789",
                         "owner_pw": "w123456789"},
                        {"name": "乙", "wm": "乙的水印", "open_pw": "o223456789",
                         "owner_pw": "w223456789"}],
                pipeline=[_attach([str(src)], watermark=True)])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    cur = outputs[iid]
    assert "甲的水印" in fitz.open(str(cur["甲"][0]))[0].get_text()
    assert "乙的水印" in fitz.open(str(cur["乙"][0]))[0].get_text()


def test_attach_copies_non_pdf_as_is(tmp_path):
    src = tmp_path / "note.txt"
    src.write_text("hello", encoding="utf-8")
    ev = _event(recips=[{"name": "甲"}],
                pipeline=[_attach([str(src)])])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    out = outputs[iid]["甲"][0]
    assert out.read_text(encoding="utf-8") == "hello"
    assert out.suffix == ".txt"


def test_attach_multiple_files(tmp_path):
    a = tmp_path / "a.txt"
    a.write_text("A", encoding="utf-8")
    b = tmp_path / "b.txt"
    b.write_text("B", encoding="utf-8")
    ev = _event(recips=[{"name": "甲"}],
                pipeline=[_attach([str(a), str(b)])])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    files = outputs[iid]["甲"]
    assert len(files) == 2
    assert {f.suffix for f in files} == {".txt"}


def test_per_group_granularity_produces_one_file_per_group(tmp_path):
    src = _make_pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲", "group": "G1", "open_pw": "o111111111",
                         "owner_pw": "w111111111"},
                        {"name": "乙", "group": "G1", "open_pw": "o222222222",
                         "owner_pw": "w222222222"}],
                pipeline=[{"kind": "attach", "granularity": "per_group",
                           "params": {"files": [str(src)]}}])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    final = outputs[iid]
    assert list(final.keys()) == ["G1"]
    assert len(final["G1"]) == 1


def test_all_granularity_produces_single_shared_file(tmp_path):
    src = _make_pdf(tmp_path / "src.pdf")
    ev = _event(recips=[{"name": "甲", "open_pw": "o111111111", "owner_pw": "w111111111"},
                        {"name": "乙", "open_pw": "o222222222", "owner_pw": "w222222222"}],
                pipeline=[{"kind": "attach", "granularity": "all",
                           "params": {"files": [str(src)], "encrypt": True}}])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    assert list(outputs[iid].keys()) == ["__all__"]
    assert len(outputs[iid]["__all__"]) == 1


def test_disabled_item_is_skipped(tmp_path):
    """停用的条目不产出，后面的条目照常。"""
    src = _make_pdf(tmp_path / "src.pdf")
    note = tmp_path / "note.txt"
    note.write_text("hello", encoding="utf-8")
    ev = _event(recips=[{"name": "甲"}],
                pipeline=[{"kind": "attach", "enabled": False,
                           "params": {"files": [str(src)]}},
                          {"kind": "attach",
                           "params": {"files": [str(note)]}}])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert not errs, errs
    assert len(outputs) == 1, "只有启用的那个条目参与执行"
    iid = ev["pipeline"][1]["id"]
    assert outputs[iid]["甲"][0].suffix == ".txt"


def test_field_level_sheet_source_joins_external_table(tmp_path):
    """来源下沉到字段级：sheet.XXX 按收件人匹配外部表。"""
    sheet = tmp_path / "data.csv"
    # 带千分位的值必须加引号，否则 CSV 会被当成两列
    sheet.write_text('Investor,Call_Amount,Due_Date\n'
                     '"甲公司","200,000","2026-11-15"\n'
                     '"乙公司","2,000,000","2026-11-15"\n', encoding="utf-8")
    smap, warns = load_sheet_data(str(sheet), "Investor")
    assert not warns
    assert smap["甲公司"]["Call_Amount"] == "200,000"

    it = normalize_item({
        "kind": "generate",
        "params": {"template_path": "C:/x/tpl.docx"},
        "mapping": {"{{Call_Amount}}": "sheet.Call_Amount"},
        "sheet": {"path": str(sheet), "key_col": "Investor", "match_col": "name"},
    })
    assert item_uses_sheet(it)
    ev = _event(recips=[{"name": "甲公司"}])
    ctx = RunContext(ev, tmp_path)
    row = ctx.sheet_for(it, ev["recipients"][0])
    assert row["Call_Amount"] == "200,000"
    # 映射表里能看到这个来源和它的预览值
    ph, src, val = item_placeholders(it, ev, ev["recipients"][0], row)[0]
    assert ph == "{{Call_Amount}}"
    assert src == "sheet.Call_Amount"
    assert val == "200,000"


def test_sheet_to_dicts_skips_empty_rows():
    rows = [["a", "b"], ["1", ""], ["", ""]]
    out = sheet_to_dicts(rows)
    assert out == [{"a": "1", "b": ""}]


def test_item_output_name_uses_filename_template(tmp_path):
    ev = _event(recips=[{"name": "甲公司"}])
    ev["options"]["filename_tpl"] = "{{原件名}}_{{LP名称}}"
    unit = dict(build_units(ev, "per_lp"))["甲公司"]
    name = item_output_name(ev, unit, "C:/docs/对账单.pdf")
    assert name == "对账单_甲公司"


def test_run_items_reports_missing_template(tmp_path):
    ev = _event(recips=[{"name": "甲"}],
                pipeline=[_generate(template=str(tmp_path / "nope.docx"))])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert errs and "不存在" in errs[0][2]


def test_run_items_reports_missing_attach_file(tmp_path):
    ev = _event(recips=[{"name": "甲"}],
                pipeline=[_attach([str(tmp_path / "nope.pdf")])])
    outputs, warns, errs = run_items(ev, tmp_path / "out")
    assert errs and "不存在" in errs[0][2]


def test_run_items_reports_empty_pipeline():
    ev = _event(recips=[{"name": "甲"}], pipeline=[])
    outputs, warns, errs = run_items(ev, Path("out"))
    assert errs and "流水线为空" in errs[0][2]


# ---------------------------------------------------------------- Word 环节
# word_batch / word_template 夹具见 conftest.py（跨模块共用）


def test_detect_placeholders(word_template):
    ph = wm.detect_placeholders(str(word_template))
    assert "{{LP_Name}}" in ph
    assert "{{Call_Amount}}" in ph
    assert "{{Due_Date}}" in ph


def test_generate_item_replaces_placeholders_and_exports_pdf(word_template, word_batch,
                                                             tmp_path):
    """生成文档条目 + 勾选「转成 PDF」。"""
    ev = _event(recips=[{"name": "甲公司", "to": "a@x.com",
                         "open_pw": "Open123456", "owner_pw": "Owner123456",
                         "fields": {"Call_Amount": "200,000",
                                    "Due_Date": "2026-11-15"}}],
                pipeline=[_generate(template=str(word_template), to_pdf=True)])
    outputs, warns, errs = run_items(ev, tmp_path / "out", word=word_batch)
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    pdf = outputs[iid]["甲公司"][0]
    assert pdf.suffix == ".pdf"
    d = fitz.open(str(pdf))
    try:
        text = d[0].get_text()
    finally:
        d.close()
    assert "甲公司" in text
    assert "200,000" in text
    assert "2026-11-15" in text
    assert "{{" not in text
    assert "Call Amount" in text      # 正文保留（格式保持）


def test_generate_item_without_pdf_keeps_docx(word_template, word_batch, tmp_path):
    ev = _event(recips=[{"name": "甲公司", "open_pw": "Open123456",
                         "owner_pw": "Owner123456",
                         "fields": {"Call_Amount": "200,000",
                                    "Due_Date": "2026-11-15"}}],
                pipeline=[_generate(template=str(word_template))])
    outputs, warns, errs = run_items(ev, tmp_path / "out", word=word_batch)
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    docx = outputs[iid]["甲公司"][0]
    assert docx.suffix == ".docx"
    assert docx.exists()


def test_generate_item_full_chain_to_encrypted_pdf(word_template, word_batch, tmp_path):
    """生成文档 → 转 PDF → 加水印 → 加密（一条目内全链路）。"""
    ev = _event(recips=[{"name": "甲公司", "wm": "甲公司水印",
                         "open_pw": "Open123456", "owner_pw": "Owner123456",
                         "fields": {"Call_Amount": "200,000",
                                    "Due_Date": "2026-11-15"}}],
                pipeline=[_generate(template=str(word_template), to_pdf=True,
                                    watermark=True, encrypt=True)])
    outputs, warns, errs = run_items(ev, tmp_path / "out", word=word_batch)
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    out = outputs[iid]["甲公司"][0]
    d = fitz.open(str(out))
    try:
        assert d.needs_pass
        assert d.authenticate("Open123456") >= 2
        text = d[0].get_text()
    finally:
        d.close()
    assert "200,000" in text
    assert "甲公司水印" in text
    assert "{{" not in text


def test_generate_item_uses_field_level_mapping(word_template, word_batch, tmp_path):
    """字段级来源：把 {{Call_Amount}} 指向收件人字段。"""
    ev = _event(recips=[{"name": "甲公司", "open_pw": "Open123456",
                         "owner_pw": "Owner123456",
                         "fields": {"Call_Amount": "200,000",
                                    "Due_Date": "2026-11-15"}}],
                pipeline=[_generate(template=str(word_template), to_pdf=True,
                                    mapping={"{{Call_Amount}}": "fields.Call_Amount"})])
    outputs, warns, errs = run_items(ev, tmp_path / "out", word=word_batch)
    assert not errs, errs
    iid = ev["pipeline"][0]["id"]
    pdf = outputs[iid]["甲公司"][0]
    d = fitz.open(str(pdf))
    try:
        text = d[0].get_text()
    finally:
        d.close()
    assert "200,000" in text


def test_two_generate_items_with_same_template_get_different_names(word_template,
                                                                   word_batch, tmp_path):
    """两个生成条目用同一个模板时，产物不能同名（附件包里会撞名）。"""
    ev = _event(recips=[{"name": "甲公司", "open_pw": "Open123456",
                         "owner_pw": "Owner123456",
                         "fields": {"Call_Amount": "200,000",
                                    "Due_Date": "2026-11-15"}}],
                pipeline=[_generate(template=str(word_template), to_pdf=True,
                                    name="季度对账单"),
                          _generate(template=str(word_template), to_pdf=True,
                                    name="催缴函")])
    outputs, warns, errs = run_items(ev, tmp_path / "out", word=word_batch)
    assert not errs, errs
    names = []
    for it in ev["pipeline"]:
        names += [p.name for p in outputs[it["id"]]["甲公司"]]
    assert len(names) == 2
    assert len(set(names)) == 2, "两个条目的产物不能同名：%s" % names
    assert "季度对账单_甲公司.pdf" in names
    assert "催缴函_甲公司.pdf" in names
    # 预览清单要和实际产物一致
    plan = preview_plan(ev)
    _k, _u, preview_names = plan[0]
    assert sorted(preview_names) == sorted(names)
