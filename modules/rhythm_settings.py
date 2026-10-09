# -*- coding: utf-8 -*-
"""拟人节流（wechatauto rhythm）的全局应用入口。

微信发送慢的根因（v1.5.x 排查）：第三方驱动 wechatauto 在每次发消息 Enter
前调用内置节流 ``rhythm.gate('send')``，默认档 ``natural`` 的冷却 cooloff 为
30–75 秒，连续批量发送第 7 条后触发突发上限，导致"内容已粘贴却等很久才回车"。

本模块把这层节流开放成应用可控：可在源码/全局设置里切换档位、或自定义
随机延迟区间，也可以完全关闭（等价于无节流）。所有对外写动作（发消息、
发文件、点赞、通话等）都走 ``gate``，所以这里的设置一次性作用于全流程。
"""

from wechatauto import rhythm


def apply_rhythm(settings):
    """按全局设置应用拟人节流。``settings`` 为 ``global_settings['rhythm']``。

    - ``enabled`` 为假或 ``profile == 'off'``：完全关闭节流
      （gap/burst/cooloff 归零、nap 倍率 1.0，等同旧版无节流速度）。
    - 否则切到对应档位（fast / natural / calm）。
    - 若配置了自定义随机延迟区间（min_delay ≤ max_delay 且 ≥0），额外用
      ``configure`` 覆盖 gap 为该随机区间，并把 burst 抬到几乎不限，
      从而消除 natural 档的冷却限速，改写动作间隔落在指定区间内且仍随机。
    """
    if not settings:
        return
    enabled = bool(settings.get("enabled", False))
    profile = str(settings.get("profile", "off")).strip()

    if not enabled or profile == "off":
        rhythm.set_profile("off")
        return

    rhythm.set_profile(profile)

    try:
        min_delay = float(settings.get("min_delay", 0.0))
        max_delay = float(settings.get("max_delay", 0.0))
    except (TypeError, ValueError):
        return
    if 0.0 <= min_delay <= max_delay:
        # 自定义延迟：gap 变成 [min,max] 的随机间隔，几乎不限次，无冷却
        rhythm.configure(
            gap=(min_delay, max_delay),
            burst=10 ** 9,
            cooloff=(0.0, 0.0),
            window=120.0,
        )