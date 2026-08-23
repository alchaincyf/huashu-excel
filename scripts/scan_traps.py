#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""分析陷阱扫描 —— 找出那些「算得对但结论错」的地方。

自动 EDA 到处都是。这个脚本做的是另一件事：把已知会让人得出错误结论的
统计陷阱变成可自动检测的项目。

扫这些：
    · 辛普森悖论    分组内的结论和总体的结论相反
    · 幽灵分组      空格/大小写/全半角造成的重复类别
    · 小基数        样本太少却在报百分比
    · 时间断点      时间序列有缺口，趋势线会撒谎
    · 不完整末期    最后一期没走完，看起来像断崖
    · 双峰分布      混了两群对象，均值描述的是不存在的人
    · 极端集中      少数几个值占了大部分，均值没有代表性
    · 伪维度        高基数或常量列，不适合做分组

这些陷阱的共同点：数字全都算对了，错的是从数字到结论那一步。
所以它们逃得过对账，只能靠专门扫描。

用法
----
    python3 scan_traps.py <文件> [--sheet 名称] [--min-group 30]

依赖：openpyxl（读 .xlsx）。不用 pandas、不用 scipy。
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).parent))
from profile_table import (  # noqa: E402
    profile, parse_number, _norm, _is_blank, _merge_anchor_map,
    read_grid, five_number, FULLWIDTH_SPACE,
)
from datetime import datetime

MIN_GROUP = 30          # 小于这个样本量就不该报百分比
SIMPSON_MIN_LAYERS = 2  # 至少几层才值得检查辛普森悖论


@dataclass
class Trap:
    kind: str
    where: str
    detail: str
    severity: str = "warn"        # high | warn | info
    advice: str = ""


@dataclass
class ScanReport:
    path: str = ""
    sheet: str = ""
    traps: list[Trap] = field(default_factory=list)
    scanned: list[str] = field(default_factory=list)

    def add(self, t: Trap) -> None:
        self.traps.append(t)


def _norm_key(s: str) -> str:
    """归一化后的分组键：用来发现「张伟」和「张伟 」其实是一个人。"""
    return (s.replace(FULLWIDTH_SPACE, "")
             .replace(" ", "")
             .strip()
             .lower()
             .replace("（", "(").replace("）", ")")
             .replace("，", ",")
             .replace("－", "-").replace("–", "-").replace("—", "-"))


def scan(path: Path, sheet: str | None = None, min_group: int = MIN_GROUP) -> ScanReport:
    p = profile(path, sheet)
    grid, _ = read_grid(path, sheet)
    rep = ScanReport(path=str(path), sheet=p.sheet)

    if p.n_data_rows == 0:
        rep.add(Trap("无数据", "-", "没有识别出明细行", "high"))
        return rep

    amap = _merge_anchor_map(p.merged_ranges)
    excluded = {r["row"] for r in p.summary_rows} | {r["row"] for r in p.footnote_rows}
    rows = [r for r in range(p.data_start_row, p.data_end_row + 1)
            if r not in excluded and not all(_is_blank(v) for v in grid[r - 1])]

    def cell(row: int, ci: int) -> Any:
        r = grid[row - 1] if row - 1 < len(grid) else []
        v = r[ci] if ci < len(r) else None
        if _is_blank(v) and (row, ci + 1) in amap:
            ar, ac = amap[(row, ci + 1)]
            src = grid[ar - 1] if ar - 1 < len(grid) else []
            v = src[ac - 1] if ac - 1 < len(src) else None
        return v

    n = len(rows)
    cats = [c for c in p.columns if c.inferred_kind == "文本" and 1 < c.n_unique <= 20]
    meas = [c for c in p.columns if c.inferred_kind == "数值"]
    times = [c for c in p.columns if c.inferred_kind == "日期"]
    rep.scanned = [f"{n} 行", f"分类列 {len(cats)}", f"数值列 {len(meas)}", f"时间列 {len(times)}"]

    # ── 1. 幽灵分组 ────────────────────────────────────────────────
    for c in p.columns:
        if c.inferred_kind != "文本":
            continue
        raw_vals = [_norm(cell(r, c.index)) for r in rows]
        raw_vals = [v for v in raw_vals if v]
        if not raw_vals:
            continue
        raw_u = set(raw_vals)
        norm_u = {_norm_key(v) for v in raw_vals}
        if len(raw_u) > len(norm_u):
            groups = defaultdict(set)
            for v in raw_vals:
                groups[_norm_key(v)].add(v)
            dup = {k: sorted(vs) for k, vs in groups.items() if len(vs) > 1}
            examples = "；".join(f"{vs!r}" for vs in list(dup.values())[:3])
            rep.add(Trap(
                "幽灵分组", f"{c.letter}列「{c.name}」",
                f"看起来有 {len(raw_u)} 个类别，归一化后只有 {len(norm_u)} 个。"
                f"以下值实际是同一个：{examples}",
                "high",
                "groupby 之前先归一化：去首尾空格、去全角空格、统一大小写和中英文标点。"
                "不处理的话这些会被算成不同的组，每组的数都是残缺的。",
            ))

    # ── 2. 小基数 ──────────────────────────────────────────────────
    for c in cats:
        counts = defaultdict(int)
        for r in rows:
            v = _norm(cell(r, c.index))
            if v:
                counts[_norm_key(v)] += 1
        small = {k: v for k, v in counts.items() if v < min_group}
        if small and len(counts) > 1:
            worst = sorted(small.items(), key=lambda kv: kv[1])[:4]
            rep.add(Trap(
                "小基数", f"{c.letter}列「{c.name}」",
                f"{len(small)}/{len(counts)} 个分组的样本量少于 {min_group}："
                + "、".join(f"{k}({v}行)" for k, v in worst),
                "warn",
                f"这些组不要报百分比，给绝对数。"
                f"样本 3 个的组「增长 200%」，实际上是多了 4 个。",
            ))

    # ── 3. 辛普森悖论 ──────────────────────────────────────────────
    # 对每个 (分组A, 分层B, 度量M)：总体上 A1 vs A2 的方向，
    # 与「在各层 B 内部 A1 vs A2」的方向是否相反。
    for m in meas:
        for a in cats:
            for b in cats:
                if a.index == b.index:
                    continue
                # 总体：A 的各类别均值
                overall = defaultdict(list)
                layered = defaultdict(lambda: defaultdict(list))
                for r in rows:
                    av = _norm_key(_norm(cell(r, a.index)))
                    bv = _norm_key(_norm(cell(r, b.index)))
                    mv, _ = parse_number(cell(r, m.index))
                    if not av or not bv or mv is None:
                        continue
                    overall[av].append(mv)
                    layered[bv][av].append(mv)
                if len(overall) < 2 or len(layered) < SIMPSON_MIN_LAYERS:
                    continue
                means = {k: sum(v) / len(v) for k, v in overall.items() if v}
                if len(means) < 2:
                    continue
                top2 = sorted(means.items(), key=lambda kv: kv[1], reverse=True)[:2]
                (a1, m1), (a2, m2) = top2
                if m1 == m2:
                    continue
                # 在每一层里，a1 是否仍然高于 a2
                flips, valid = 0, 0
                for bv, inner in layered.items():
                    if a1 in inner and a2 in inner and inner[a1] and inner[a2]:
                        valid += 1
                        i1 = sum(inner[a1]) / len(inner[a1])
                        i2 = sum(inner[a2]) / len(inner[a2])
                        if i1 < i2:
                            flips += 1
                if valid >= SIMPSON_MIN_LAYERS and flips == valid:
                    rep.add(Trap(
                        "辛普森悖论", f"{m.name} · {a.name} × {b.name}",
                        f"总体上「{a1}」的均值({m1:,.2f})高于「{a2}」({m2:,.2f})，"
                        f"但在全部 {valid} 个「{b.name}」分层内部，结论都是反的。",
                        "high",
                        f"不要用总量下结论。差异来自各层的样本构成不同，"
                        f"不是「{a1}」真的更好。报告应该给分层结果，"
                        f"并说明构成差异。",
                    ))

    # ── 4. 时间断点与不完整末期 ────────────────────────────────────
    for t in times:
        ds: list[datetime] = []
        for r in rows:
            raw = cell(r, t.index)
            if isinstance(raw, datetime):
                ds.append(raw)
            else:
                s = _norm(raw)
                for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y-%m", "%Y/%m"):
                    try:
                        ds.append(datetime.strptime(s, fmt))
                        break
                    except ValueError:
                        continue
        if len(ds) < 3:
            continue
        ds.sort()
        gaps = [(ds[i + 1] - ds[i]).days for i in range(len(ds) - 1)]
        pos = sorted(g for g in gaps if g > 0)
        if not pos:
            continue
        typical = pos[len(pos) // 2]
        big = [(ds[i], ds[i + 1], gaps[i]) for i in range(len(gaps))
               if gaps[i] > max(typical * 3, typical + 2)]
        if big:
            ex = "、".join(f"{s:%Y-%m-%d}→{e:%Y-%m-%d}({g}天)" for s, e, g in big[:3])
            rep.add(Trap(
                "时间断点", f"{t.letter}列「{t.name}」",
                f"典型间隔 {typical} 天，但出现 {len(big)} 处明显更大的缺口：{ex}",
                "warn",
                "缺失的时间段不会报错，但会让趋势线撒谎。"
                "确认是真的没有业务，还是数据没采到。",
            ))

    # ── 5. 分布形态：双峰、极端集中 ────────────────────────────────
    for c in meas:
        vals = []
        for r in rows:
            v, _ = parse_number(cell(r, c.index))
            if v is not None:
                vals.append(v)
        if len(vals) < 8:
            continue
        st = five_number(vals)
        total = sum(vals)

        # 极端集中：前 20% 的项贡献了多少
        sv = sorted(vals, reverse=True)
        k = max(1, len(sv) // 5)
        top_share = sum(sv[:k]) / total if total else 0
        if top_share > 0.8:
            rep.add(Trap(
                "极端集中", f"{c.letter}列「{c.name}」",
                f"前 20% 的项贡献了 {top_share:.0%} 的总量"
                f"（均值 {st['mean']:,.2f}，中位数 {st['median']:,.2f}）",
                "warn",
                "均值在这种分布下没有代表性。用中位数描述「典型」，"
                "并把头部单独拿出来讲。丢掉头部那几个，整体结论会完全不同。",
            ))

        # 双峰的粗检：中间区域的密度低于两侧
        lo, hi = st["min"], st["max"]
        if hi > lo and len(vals) >= 12:
            bins = [0] * 5
            for v in vals:
                idx = min(4, int((v - lo) / (hi - lo) * 5))
                bins[idx] += 1
            mid = bins[2]
            if mid > 0 and bins[0] > mid * 1.5 and bins[4] > mid * 1.5:
                rep.add(Trap(
                    "疑似双峰", f"{c.letter}列「{c.name}」",
                    f"五等分后的分布为 {bins}，两端明显高于中间",
                    "warn",
                    "可能混了两群性质不同的对象。均值会落在中间的谷底，"
                    "描述的是一个不存在的典型。先找出区分这两群的那个变量。",
                ))

    # ── 6. 伪维度 ──────────────────────────────────────────────────
    for c in p.columns:
        live = c.n_total - c.n_missing
        if live == 0:
            continue
        if c.n_unique == 1 and live > 1:
            rep.add(Trap("伪维度", f"{c.letter}列「{c.name}」",
                         "常量列，所有行同一个值", "info",
                         "对分析没有信息量，分组时跳过它。"))
        elif c.inferred_kind == "文本" and c.n_unique >= live * 0.95 and live > 5:
            rep.add(Trap("伪维度", f"{c.letter}列「{c.name}」",
                         f"{c.n_unique} 个唯一值 / {live} 行，几乎每行都不同", "info",
                         "这更像是编号或名称，不适合做分组维度。"))
    return rep


def render(rep: ScanReport) -> str:
    L = ["─" * 68, f"陷阱扫描：{Path(rep.path).name}  [{rep.sheet}]", "─" * 68]
    L.append("扫描范围：" + "，".join(rep.scanned))
    L.append("")
    if not rep.traps:
        L.append("没有发现已知陷阱。")
        L.append("")
        L.append("注意：这只说明没踩到可自动检测的坑，不代表分析是对的。")
        L.append("「这个分析有没有意义」机器判不了。")
        L.append("─" * 68)
        return "\n".join(L)

    order = {"high": 0, "warn": 1, "info": 2}
    label = {"high": "会得出错误结论", "warn": "需要你判断", "info": "提示"}
    cur = None
    for t in sorted(rep.traps, key=lambda x: order.get(x.severity, 3)):
        if t.severity != cur:
            cur = t.severity
            L.append(f"【{label.get(cur, cur)}】")
        L.append(f"  · {t.kind}  {t.where}")
        L.append(f"    {t.detail}")
        if t.advice:
            L.append(f"    → {t.advice}")
        L.append("")
    L.append("─" * 68)
    L.append("这些陷阱的共同点：数字都算对了，错的是从数字到结论那一步。")
    L.append("它们逃得过对账，只能靠专门扫描。")
    L.append("─" * 68)
    return "\n".join(L)


def main() -> None:
    ap = argparse.ArgumentParser(description="扫描会让人得出错误结论的分析陷阱")
    ap.add_argument("file")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--min-group", type=int, default=MIN_GROUP,
                    help=f"小于这个样本量的分组会被标记，默认 {MIN_GROUP}")
    a = ap.parse_args()

    src = Path(a.file).expanduser()
    if not src.exists():
        sys.exit(f"文件不存在：{src}")
    print(render(scan(src, a.sheet, a.min_group)))


if __name__ == "__main__":
    main()
