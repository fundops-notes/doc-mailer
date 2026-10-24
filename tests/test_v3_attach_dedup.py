# -*- coding: utf-8 -*-
"""合并组发送语义 + 附件内容级去重（哈希）的回归测试。

用户拍板的语义（2026-10-24）：
- 有合并组 → 同组一封（共用模板、附件并入、收件人去重）
- 没有组 → 各发各的；不同组 → 不同邮件
- 附件去重两层：路径 → 内容指纹（大小+SHA256），同内容不同名只保留一份
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import pytest

import lp_v3


def _ev(recips, send_mode="per_group"):
    ev = lp_v3.new_event("t")
    ev["options"]["send_mode"] = send_mode
    out = []
    for r in recips:
        rec = lp_v3.normalize_recipient(r)
        out.append(rec)
    ev["recipients"] = out
    return ev


class TestGroupSemantics:
    def test_per_group_hybrid_no_group_gets_own_email(self):
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1"},
            {"name": "乙", "to": "b@x.cn", "group": "G1"},
            {"name": "丙", "to": "c@x.cn", "group": ""},
            {"name": "丁", "to": "d@x.cn", "group": "G2"},
            {"name": "戊", "to": "e@x.cn", "group": "G2"},
            {"name": "己", "to": "f@x.cn", "group": ""},
        ])
        units = lp_v3.draft_units(ev)
        members = [tuple(r["name"] for r in u["recipients"]) for _, u in units]
        assert ("甲", "乙") in members          # G1 合并
        assert ("丁", "戊") in members          # G2 合并
        assert ("丙",) in members and ("己",) in members  # 无组各自一封
        assert len(units) == 4

    def test_merged_addresses_dedup(self):
        ev = _ev([
            {"name": "甲", "to": "a@x.cn;shared@x.cn", "cc": "cc1@x.cn"},
            {"name": "乙", "to": "b@x.cn;shared@x.cn", "cc": "cc1@x.cn;cc2@x.cn"},
        ])
        to = lp_v3._merge_addresses([r for r in ev["recipients"]], "to")
        cc = lp_v3._merge_addresses([r for r in ev["recipients"]], "cc")
        assert to == ["a@x.cn", "shared@x.cn", "b@x.cn"]   # 去重且保序
        assert cc == ["cc1@x.cn", "cc2@x.cn"]

    def test_per_lp_still_individual(self):
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1"},
            {"name": "乙", "to": "b@x.cn", "group": "G1"},
        ], send_mode="per_lp")
        assert len(lp_v3.draft_units(ev)) == 2


class TestHashDedup:
    def test_identical_content_different_names_keeps_one(self, tmp_path):
        f1 = tmp_path / "通知_甲.pdf"
        f2 = tmp_path / "通知_乙.pdf"
        payload = b"%PDF-1.4 same-bytes" * 10
        f1.write_bytes(payload)
        f2.write_bytes(payload)
        ev = _ev([{"name": "甲", "to": "a@x.cn", "group": "G1"}])
        unit = lp_v3.draft_units(ev)[0][1]
        dup = []
        files, missing = lp_v3.assemble_attachments(
            ev, unit, [str(f1), str(f2)], dup_log=dup)
        assert len(files) == 1 and missing == []
        assert files[0] == f1                       # 保留首个
        assert len(dup) == 1 and dup[0][1] == f2    # 记录被丢弃者

    def test_different_content_both_kept(self, tmp_path):
        f1 = tmp_path / "a.pdf"; f1.write_bytes(b"content-A")
        f2 = tmp_path / "b.pdf"; f2.write_bytes(b"content-B")
        ev = _ev([{"name": "甲", "to": "a@x.cn"}])
        unit = lp_v3.draft_units(ev)[0][1]
        files, _ = lp_v3.assemble_attachments(ev, unit, [str(f1), str(f2)])
        assert len(files) == 2

    def test_same_path_twice_one_attachment(self, tmp_path):
        f1 = tmp_path / "same.pdf"; f1.write_bytes(b"x" * 32)
        ev = _ev([{"name": "甲", "to": "a@x.cn"}])
        unit = lp_v3.draft_units(ev)[0][1]
        files, _ = lp_v3.assemble_attachments(ev, unit,
                                              [str(f1), str(f1)])
        assert len(files) == 1

    def test_static_and_recipient_dedup_across_sources(self, tmp_path):
        """事件静态附件与收件人专属附件同内容 → 合并组邮件里只挂一份"""
        shared = tmp_path / "说明书.pdf"
        shared.write_bytes(b"manual-bytes" * 5)
        copy_ = tmp_path / "说明书_副本.pdf"
        copy_.write_bytes(b"manual-bytes" * 5)      # 同内容不同名
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1",
             "attachments": [str(copy_)]},
            {"name": "乙", "to": "b@x.cn", "group": "G1"},
        ])
        ev["attachments"] = [{"name": "说明书", "path": str(shared),
                              "scope": "all"}]
        unit = lp_v3.draft_units(ev)[0][1]
        dup = []
        files, _ = lp_v3.assemble_attachments(ev, unit, [], dup_log=dup)
        assert len(files) == 1
        assert files[0] == shared                   # 先到的优先

    def test_backward_compatible_signature(self, tmp_path):
        """老调用不传 dup_log 也照常工作"""
        f1 = tmp_path / "x.pdf"; f1.write_bytes(b"z" * 16)
        ev = _ev([{"name": "甲", "to": "a@x.cn"}])
        unit = lp_v3.draft_units(ev)[0][1]
        files, missing = lp_v3.assemble_attachments(ev, unit, [str(f1)])
        assert len(files) == 1 and missing == []


class TestFileDigest:
    def test_digest_stable_and_sensitive(self, tmp_path):
        a = tmp_path / "a.bin"; a.write_bytes(b"hello")
        b = tmp_path / "b.bin"; b.write_bytes(b"hello")
        c = tmp_path / "c.bin"; c.write_bytes(b"world")
        assert lp_v3._file_digest(a) == lp_v3._file_digest(b)
        assert lp_v3._file_digest(a) != lp_v3._file_digest(c)

    def test_digest_none_on_missing(self, tmp_path):
        assert lp_v3._file_digest(tmp_path / "nope.bin") is None
