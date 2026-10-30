# -*- coding: utf-8 -*-
"""v3.9.17 编码加固回归测试（2026-10-30 NAV乱码事件）。

覆盖 create_outlook_draft 的代码页显式声明：
  · 两项代码页属性（PR_MESSAGE_CODEPAGE / PR_INTERNET_CPID）逐项写入
  · 单项失败时：不抛出、其余项照常写入、结果标记失败
  · 全部失败时：不抛出、结果标记失败
  · Save 后读回验证持久化；读回不一致时结果标记未持久化
  · hardening_result 未传入时不影响原调用方

使用假 Outlook 对象（无需真实 Outlook / win32com）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import lp_v3

PROPS = {prop: (value, key) for prop, value, key in lp_v3._CODEPAGE_HARDENING_PROPS}


class FakePropertyAccessor:
    """模拟 Outlook PropertyAccessor：记录写入，按需注入故障，支持读回。"""

    def __init__(self, fail_on=()):
        self.fail_on = set(fail_on)          # 注入故障的属性 tag
        self.values = {}                     # tag -> value
        self.set_calls = []                  # [(tag, value)]

    def SetProperty(self, prop, value):
        if prop in self.fail_on:
            raise OSError("simulated SetProperty failure: %s" % prop)
        self.set_calls.append((prop, value))
        self.values[prop] = value

    def GetProperty(self, prop):
        if prop in self.values:
            return self.values[prop]
        raise OSError("property not set: %s" % prop)


class FakeMail:
    """模拟 MailItem：只提供加固逻辑所需的 PropertyAccessor。"""

    def __init__(self, fail_on=()):
        self._pa = FakePropertyAccessor(fail_on)
        self.saved = 0

    @property
    def PropertyAccessor(self):
        return self._pa

    def Save(self):
        self.saved += 1


def _run(mail):
    result = {}
    lp_v3._apply_codepage_hardening(mail)
    mail.Save()
    result["hardening"] = lp_v3._readback_codepage_hardening(mail)
    # 与 create_outlook_draft 的持久化判定保持一致
    result["persisted"] = all(
        result["hardening"].get(key) == value
        for prop, value, key in lp_v3._CODEPAGE_HARDENING_PROPS
    )
    return result


def test_hardening_sets_both_properties_and_persists():
    mail = FakeMail()
    result = lp_v3._apply_codepage_hardening(mail)
    assert result["ok"] is True
    assert result["errors"] == []
    assert len(result["items"]) == 2
    assert all(e["set"] for e in result["items"])
    # 写入的值与声明一致
    for prop, value, key in lp_v3._CODEPAGE_HARDENING_PROPS:
        assert mail._pa.values[prop] == value


def test_hardening_survives_single_property_failure():
    # 第一项写入失败：不抛出、第二项照常写入、结果标记失败
    tags = [prop for prop, _, _ in lp_v3._CODEPAGE_HARDENING_PROPS]
    mail = FakeMail(fail_on={tags[0]})
    result = lp_v3._apply_codepage_hardening(mail)
    assert result["ok"] is False
    assert len(result["errors"]) == 1
    assert result["items"][0]["set"] is False
    assert result["items"][0]["error"]
    assert result["items"][1]["set"] is True
    # 第二项的值仍然写进去了
    assert mail._pa.values[tags[1]] == lp_v3._CODEPAGE_HARDENING_PROPS[1][1]


def test_hardening_survives_total_failure():
    tags = [prop for prop, _, _ in lp_v3._CODEPAGE_HARDENING_PROPS]
    mail = FakeMail(fail_on=set(tags))
    result = lp_v3._apply_codepage_hardening(mail)
    assert result["ok"] is False
    assert all(e.get("set") is False for e in result["items"])
    assert len(result["errors"]) == 2


def test_readback_reports_persisted_values():
    mail = FakeMail()
    lp_v3._apply_codepage_hardening(mail)
    readback = lp_v3._readback_codepage_hardening(mail)
    for prop, value, key in lp_v3._CODEPAGE_HARDENING_PROPS:
        assert readback[key] == value


def test_readback_reports_missing_values():
    mail = FakeMail()
    readback = lp_v3._readback_codepage_hardening(mail)  # 未写入任何值
    for prop, value, key in lp_v3._CODEPAGE_HARDENING_PROPS:
        assert readback[key] is None


def test_hardening_result_pipeline_matches_create_outlook_draft_semantics():
    mail = FakeMail()
    hardening = lp_v3._apply_codepage_hardening(mail)
    mail.Save()
    hardening["readback"] = lp_v3._readback_codepage_hardening(mail)
    hardening["persisted"] = all(
        hardening.get("readback", {}).get(key) == value
        for _, value, key in lp_v3._CODEPAGE_HARDENING_PROPS
    )
    hardening["ok"] = bool(hardening.get("ok")) and hardening["persisted"]
    assert hardening["ok"] is True
    assert hardening["persisted"] is True
