# -*- coding: utf-8 -*-
"""合并组发送语义 + 附件内容指纹提示（只提示不去重）的回归测试。

用户拍板的语义（2026-10-24 第二轮）：
- 有合并组 → 同组一封（共用模板、附件并入、收件人去重）
- 没有组 → 各发各的；不同组 → 不同邮件
- 附件去重仅按路径；**同封邮件内不同路径但内容（大小+SHA256）相同的，
  两份全部保留，只在检查/预检/运行日志里特别提示**，不改路径、不去重
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import lp_v3


def _ev(recips, send_mode="per_group"):
    ev = lp_v3.new_event("t")
    ev["options"]["send_mode"] = send_mode
    ev["recipients"] = [lp_v3.normalize_recipient(r) for r in recips]
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
        assert ("甲", "乙") in members
        assert ("丁", "戊") in members
        assert ("丙",) in members and ("己",) in members
        assert len(units) == 4

    def test_merged_addresses_dedup(self):
        ev = _ev([
            {"name": "甲", "to": "a@x.cn;shared@x.cn", "cc": "cc1@x.cn"},
            {"name": "乙", "to": "b@x.cn;shared@x.cn", "cc": "cc1@x.cn;cc2@x.cn"},
        ])
        to = lp_v3._merge_addresses([r for r in ev["recipients"]], "to")
        cc = lp_v3._merge_addresses([r for r in ev["recipients"]], "cc")
        assert to == ["a@x.cn", "shared@x.cn", "b@x.cn"]
        assert cc == ["cc1@x.cn", "cc2@x.cn"]

    def test_per_lp_still_individual(self):
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1"},
            {"name": "乙", "to": "b@x.cn", "group": "G1"},
        ], send_mode="per_lp")
        assert len(lp_v3.draft_units(ev)) == 2


class TestFingerprintWarnOnly:
    """同内容不同路径 → 全部保留 + dup_log 记录；同路径 → 只挂一份。"""

    def test_identical_content_different_names_both_kept(self, tmp_path):
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
        assert len(files) == 2                       # 两份都保留
        assert missing == []
        assert len(dup) == 1 and dup[0] == (f1, f2)  # 记录提示对

    def test_same_path_twice_one_attachment(self, tmp_path):
        f1 = tmp_path / "same.pdf"; f1.write_bytes(b"x" * 32)
        ev = _ev([{"name": "甲", "to": "a@x.cn"}])
        unit = lp_v3.draft_units(ev)[0][1]
        files, _ = lp_v3.assemble_attachments(ev, unit, [str(f1), str(f1)])
        assert len(files) == 1                       # 路径去重保留

    def test_different_content_both_kept_no_warning(self, tmp_path):
        f1 = tmp_path / "a.pdf"; f1.write_bytes(b"content-A")
        f2 = tmp_path / "b.pdf"; f2.write_bytes(b"content-B")
        ev = _ev([{"name": "甲", "to": "a@x.cn"}])
        unit = lp_v3.draft_units(ev)[0][1]
        dup = []
        files, _ = lp_v3.assemble_attachments(ev, unit, [str(f1), str(f2)],
                                              dup_log=dup)
        assert len(files) == 2 and dup == []

    def test_backward_compatible_signature(self, tmp_path):
        f1 = tmp_path / "x.pdf"; f1.write_bytes(b"z" * 16)
        ev = _ev([{"name": "甲", "to": "a@x.cn"}])
        unit = lp_v3.draft_units(ev)[0][1]
        files, missing = lp_v3.assemble_attachments(ev, unit, [str(f1)])
        assert len(files) == 1 and missing == []


class TestFingerprintWarnings:
    """attachment_fingerprint_warnings：按邮件单元提示，跨来源识别。"""

    def test_warns_same_email_only(self, tmp_path):
        """同封邮件内重复才提示；分属两封的不提示。"""
        dup_payload = b"DUP" * 100
        f_a = tmp_path / "甲件.pdf"; f_a.write_bytes(dup_payload)
        f_b = tmp_path / "乙件.pdf"; f_b.write_bytes(dup_payload)
        f_c = tmp_path / "丙件.pdf"; f_c.write_bytes(dup_payload)
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1",
             "attachments": [str(f_a)]},
            {"name": "乙", "to": "b@x.cn", "group": "G1",
             "attachments": [str(f_b)]},
            {"name": "丙", "to": "c@x.cn", "group": "",
             "attachments": [str(f_c)]},
        ])
        warns = lp_v3.attachment_fingerprint_warnings(ev)
        # G1 那封里 甲件/乙件 重复 → 提示；丙独封 → 不提示
        assert len(warns) == 1 and "甲件" in warns[0] and "乙件" in warns[0]

    def test_no_warning_for_same_path(self, tmp_path):
        f = tmp_path / "only.pdf"; f.write_bytes(b"S" * 8)
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1"},
            {"name": "乙", "to": "b@x.cn", "group": "G1"},
        ])
        ev["attachments"] = [{"name": "共用", "path": str(f), "scope": "all"}]
        assert lp_v3.attachment_fingerprint_warnings(ev) == []

    def test_static_and_recipient_cross_source(self, tmp_path):
        shared = tmp_path / "说明书.pdf"
        shared.write_bytes(b"manual-bytes" * 5)
        copy_ = tmp_path / "说明书_副本.pdf"
        copy_.write_bytes(b"manual-bytes" * 5)
        ev = _ev([
            {"name": "甲", "to": "a@x.cn", "group": "G1",
             "attachments": [str(copy_)]},
        ])
        ev["attachments"] = [{"name": "说明书", "path": str(shared),
                              "scope": "all"}]
        warns = lp_v3.attachment_fingerprint_warnings(ev)
        assert len(warns) == 1 and "说明书" in warns[0]

    def test_preflight_includes_fingerprint_warnings(self, tmp_path):
        dup_payload = b"P" * 64
        f1 = tmp_path / "x1.pdf"; f1.write_bytes(dup_payload)
        f2 = tmp_path / "x2.pdf"; f2.write_bytes(dup_payload)
        ev = _ev([{"name": "甲", "to": "a@x.cn",
                   "attachments": [str(f1), str(f2)]}])
        _errors, warnings, _info = lp_v3.preflight(ev)
        assert any("x1.pdf" in w and "x2.pdf" in w for w in warnings)


class TestFileDigest:
    def test_digest_stable_and_sensitive(self, tmp_path):
        a = tmp_path / "a.bin"; a.write_bytes(b"hello")
        b = tmp_path / "b.bin"; b.write_bytes(b"hello")
        c = tmp_path / "c.bin"; c.write_bytes(b"world")
        assert lp_v3._file_digest(a) == lp_v3._file_digest(b)
        assert lp_v3._file_digest(a) != lp_v3._file_digest(c)

    def test_digest_none_on_missing(self, tmp_path):
        assert lp_v3._file_digest(tmp_path / "nope.bin") is None
