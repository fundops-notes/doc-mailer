"""Long names, Windows-compatible paths and collision protection."""
from pathlib import Path
import sys
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import lp_v3


def test_long_fields_and_template_are_not_truncated():
    rec = {"name": "N" * 80, "fields": {"custom": "C" * 70}}
    unit = {"rep": rec}
    event = {"options": {}, "recipients": [rec]}
    name = lp_v3.item_output_name(event, unit, "source.pdf", filename_tpl="{{收件人名称}}_{{custom}}")
    assert name == "N" * 80 + "_" + "C" * 70


def test_long_common_prefixes_do_not_collide(tmp_path):
    source = tmp_path / "src.pdf"
    source.write_bytes(b"content")
    recipients = [{"name": "A" * 95 + "1"}, {"name": "A" * 95 + "2"}]
    event = {"options": {}, "recipients": recipients}
    item = lp_v3.normalize_item({"id": "step", "kind": "attach", "params": {
        "files": [str(source)], "filename_tpl": "{{收件人名称}}"}})
    context = lp_v3.RunContext(event, tmp_path / "out")
    results = [lp_v3.execute_item(item, {"rep": r, "recipients": [r]}, event, context)[0][0] for r in recipients]
    assert results[0] != results[1]
    assert [p.read_bytes() for p in results] == [b"content", b"content"]


def test_duplicate_names_get_distinct_output(tmp_path):
    ctx = lp_v3.RunContext({"options": {}}, tmp_path)
    first = ctx.reserve_output(tmp_path, "X" * 100, ".pdf")
    second = ctx.reserve_output(tmp_path, "X" * 100, ".pdf")
    assert first != second and second.stem.endswith("_2")


def test_overlong_path_fails_before_copy(tmp_path):
    source = tmp_path / "source.pdf"
    source.write_bytes(b"source")
    recipient = {"name": "Alice"}
    event = {"options": {}, "recipients": [recipient]}
    item = lp_v3.normalize_item({"id": "step", "kind": "attach", "params": {
        "files": [str(source)], "filename_tpl": "X" * 240}})
    with pytest.raises(ValueError, match="输出完整路径超过"):
        lp_v3.execute_item(item, {"rep": recipient}, event, lp_v3.RunContext(event, tmp_path / "out"))
    assert not list((tmp_path / "out").rglob("*.pdf"))


def test_temporary_pdf_path_and_utf16_component_are_checked(tmp_path):
    with pytest.raises(ValueError, match="文件名超过"):
        lp_v3._validate_output_path(tmp_path / (("😀" * 126) + ".pdf"))
    parent = Path("C:/") / ("a" * 120) / ("b" * 110)
    lp_v3._validate_output_path(parent / ("c" * 18 + ".pdf"))
    with pytest.raises(ValueError, match="完整路径超过"):
        lp_v3.RunContext({"options": {}}, parent).reserve_output(parent, "c" * 18, ".pdf", pdf_temp=True)
