#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数字对账 —— 交付之前，用表自己的数据证明你没算错。

核心思路
--------
中文办公表格几乎都自带「合计」「小计」行。所有工具都把它们当噪音过滤掉，
但它们其实是**免费的校验和**：那是原表作者用 Excel 公式算出来的真值。

拿清洗后的明细自己求和，去对那个合计。对得上，你的清洗是对的；对不上，
要么你的清洗错了，要么原表错了 —— 两种情况都必须报出来，不能沉默。

这对应 ICAEW《Financial Modelling Code》的两条：
    Include a master check   ——  一个总检查，任何一项失败就报警
    Build traceable references —— 每个数字都能追回它的来源

同时执行 Twyman's Law：任何看起来有趣或异常的数字，通常都是错的。
本脚本把「异常」量化成可自动判定的检查项，不靠人去凭感觉发现。

用法
----
    python3 verify_numbers.py <文件> [--sheet 名称] [--json]

    # 顺便核对一个你自己算出来的数
    python3 verify_numbers.py <文件> --claim "1月:12387317"

退出码：0 = 全部通过；1 = 有对不上的地方（可直接用在流水线里当闸门）
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).parent))
from profile_table import (  # noqa: E402
    profile, parse_number, _norm, _is_blank,
    TableProfile, read_grid, _merge_anchor_map,
)

# 汇总行是分层的：「华东小计」管一个分组，「合计」管整张表。
# 把它们当成平级的连续区段来划分覆盖范围，遇到两级汇总必然算错。
GROUP_LEVEL_WORDS = ("小计", "subtotal")
TOTAL_LEVEL_WORDS = ("合计", "总计", "总和", "累计", "汇总", "共计", "总额",
                     "total", "grand total", "sum")
RATIO_WORDS = ("占比", "比例", "同比", "环比", "增长", "率", "rate", "%")


def summary_level(label: str) -> str:
    """判断汇总行管辖的层级：group（分组小计）/ total（全表合计）/ unknown。"""
    t = label.strip().lower()
    if any(w in t for w in GROUP_LEVEL_WORDS):
        return "group"
    if any(w in t for w in TOTAL_LEVEL_WORDS):
        return "total"
    return "unknown"

# 相对误差容忍度：小于这个值视作「因四舍五入产生的差异」而非真错
REL_TOL = 1e-6
# 量级怀疑阈值：列名声明的单位与实际数值量级差这么多倍就报警
MAGNITUDE_FACTOR = 100


@dataclass
class Check:
    name: str
    passed: bool
    detail: str
    severity: str = "error"          # error | warn | info
    evidence: dict[str, Any] = field(default_factory=dict)


@dataclass
class VerifyReport:
    path: str = ""
    sheet: str = ""
    checks: list[Check] = field(default_factory=list)
    master_check: bool = True
    caliber: dict[str, Any] = field(default_factory=dict)

    def add(self, c: Check) -> None:
        self.checks.append(c)
        if not c.passed and c.severity == "error":
            self.master_check = False


# ── 单位量级 ──────────────────────────────────────────────────────────
UNIT_SCALE = {
    "亿元": 1e8, "亿": 1e8,
    "万元": 1e4, "万": 1e4,
    "千元": 1e3,
    "元": 1.0,
}


def declared_unit(col_name: str) -> tuple[str, float] | None:
    """从列名里读出声明的单位，如「上半年合计(万元)」→ ('万元', 1e4)。"""
    for u in sorted(UNIT_SCALE, key=len, reverse=True):
        if u in col_name:
            return u, UNIT_SCALE[u]
    return None


# ── 主流程 ────────────────────────────────────────────────────────────
def verify(path: Path, sheet: str | None = None,
           claims: dict[str, float] | None = None) -> VerifyReport:
    p: TableProfile = profile(path, sheet)
    grid, _meta = read_grid(path, sheet)
    rep = VerifyReport(path=str(path), sheet=p.sheet)

    width = len(p.columns)
    if width == 0 or p.n_data_rows == 0:
        rep.add(Check("可解析性", False, "没有识别出明细数据，无法对账"))
        return rep

    # 明细行号（1-based），已排除汇总/脚注
    excluded = {r["row"] for r in p.summary_rows} | {r["row"] for r in p.footnote_rows}
    detail_rows = [r for r in range(p.data_start_row, p.data_end_row + 1)
                   if r not in excluded and not all(_is_blank(v) for v in grid[r - 1])]
    dup_rows = {d["row"] for d in p.duplicate_rows}

    # ── 口径声明：先把话说清楚，再谈数字 ──────────────────────────
    rep.caliber = {
        "明细行范围": f"第 {p.data_start_row}-{p.data_end_row} 行",
        "明细行数": len(detail_rows),
        "已排除行": sorted(excluded) or "无",
        "重复行": sorted(dup_rows) or "无",
        "去重后行数": len(detail_rows) - len(dup_rows),
        "表头行": p.header_rows,
        "说明": "以下所有数字均基于「明细行、未去重」口径。若业务上应当去重，"
                "结果会变化，见「去重影响」检查项。",
    }

    # ── 检查 1：行数守恒 ────────────────────────────────────────────
    span = p.data_end_row - p.data_start_row + 1
    blanks = sum(1 for r in range(p.data_start_row, p.data_end_row + 1)
                 if all(_is_blank(v) for v in grid[r - 1]))
    accounted = len(detail_rows) + len([r for r in excluded
                                        if p.data_start_row <= r <= p.data_end_row]) + blanks
    rep.add(Check(
        "行数守恒",
        accounted == span,
        f"数据区跨 {span} 行 = 明细 {len(detail_rows)} + 已排除 "
        f"{accounted - len(detail_rows) - blanks} + 空行 {blanks}"
        + ("" if accounted == span else f"  ← 差 {span - accounted} 行没有归属，说明分类逻辑漏了情况"),
        evidence={"span": span, "detail": len(detail_rows), "blank": blanks},
    ))

    # ── 合并单元格 ffill：分组标签列的空值要先恢复，否则没法按组匹配 ──
    amap = _merge_anchor_map(p.merged_ranges)

    def cell(row: int, ci: int) -> Any:
        """取单元格值，落在合并区内的取其左上角的值。"""
        r = grid[row - 1] if row - 1 < len(grid) else []
        v = r[ci] if ci < len(r) else None
        if _is_blank(v) and (row, ci + 1) in amap:
            ar, ac = amap[(row, ci + 1)]
            src = grid[ar - 1] if ar - 1 < len(grid) else []
            v = src[ac - 1] if ac - 1 < len(src) else None
        return v

    def num_at(row: int, ci: int) -> float | None:
        v, _ = parse_number(cell(row, ci))
        return v

    # ── 检查 2：与表内汇总行对账（核心）────────────────────────────
    summary_lookup = {r["row"]: r for r in p.summary_rows}
    text_cols = [c for c in p.columns if c.inferred_kind == "文本"]
    checked_any = False

    def resolve_scope(srow: int, label: str) -> tuple[list[int], str] | None:
        """这个汇总行到底管哪些明细行？定不下来就返回 None，不猜。"""
        level = summary_level(label)
        if level == "total":
            return detail_rows, "全表明细"
        if level == "group":
            # 从「华东小计」剥出「华东」，再去各文本列里找这个分组
            name = label
            for w in GROUP_LEVEL_WORDS:
                name = name.replace(w, "").replace(w.upper(), "")
            name = name.strip("　 -—_:：()（）")
            if name:
                for tc in text_cols:
                    matched = [r for r in detail_rows
                               if _norm(cell(r, tc.index)) == name]
                    if matched:
                        return matched, f"{tc.letter}列「{tc.name}」== {name}"
        return None

    for srow, sinfo in sorted(summary_lookup.items()):
        label = sinfo.get("label", "")
        if any(w in label for w in RATIO_WORDS):
            continue                      # 比率行不是求和，不参与对账
        scope = resolve_scope(srow, label)
        if scope is None:
            rep.add(Check(
                f"对账 · 第{srow}行「{label}」",
                True,
                "定不下来这一行管哪些明细，跳过对账。"
                "（宁可不对，也不用推断出来的范围去对——那会给出一个错的结论。）"
                "请人工确认它的覆盖范围。",
                severity="warn",
            ))
            continue
        covered, basis = scope

        for col in p.columns:
            if col.inferred_kind != "数值":
                continue
            claimed = num_at(srow, col.index)
            if claimed is None:
                continue
            vals = [num_at(r, col.index) for r in covered]
            actual = sum(v for v in vals if v is not None)

            if claimed == 0:
                rel = 0.0 if abs(actual) < REL_TOL else float("inf")
            else:
                rel = abs(actual - claimed) / abs(claimed)
            ok = rel <= REL_TOL
            checked_any = True

            rep.add(Check(
                f"对账 · {col.letter}列「{col.name}」 vs 第{srow}行「{label}」",
                ok,
                (f"表内写 {claimed:,.2f}，按「{basis}」取 {len(covered)} 行"
                 f"算得 {actual:,.2f}"
                 + ("  ✓ 一致" if ok else
                    f"  ✗ 差 {actual - claimed:,.2f}（{rel:.2%}）")),
                severity="error" if not ok else "info",
                evidence={"claimed": claimed, "actual": actual, "basis": basis,
                          "rows": covered, "cell": f"{col.letter}{srow}"},
            ))

    if not checked_any:
        rep.add(Check(
            "对账基准", True,
            "表内没有可用的汇总行做交叉验证。这意味着结果没有独立的校验来源 —— "
            "交付前请人工抽查至少 3 个数字回溯到源单元格。",
            severity="warn",
        ))

    # ── 检查 3：去重的影响有多大 ───────────────────────────────────
    if dup_rows:
        for col in p.columns:
            if col.inferred_kind != "数值":
                continue
            ci = col.index

            def v(r: int) -> float:
                x, _ = parse_number(grid[r - 1][ci]) if ci < len(grid[r - 1]) else (None, [])
                return x or 0.0

            with_dup = sum(v(r) for r in detail_rows)
            without = sum(v(r) for r in detail_rows if r not in dup_rows)
            if with_dup == 0:
                continue
            delta = (with_dup - without) / abs(with_dup)
            rep.add(Check(
                f"去重影响 · {col.letter}列「{col.name}」",
                True,
                f"含重复 {with_dup:,.2f}，去重后 {without:,.2f}，"
                f"相差 {delta:.2%} —— 这两个数都对，取哪个取决于业务口径。"
                f"必须明确声明用的是哪个。",
                severity="warn",
                evidence={"with_dup": with_dup, "dedup": without},
            ))

    # ── 检查 4：单位量级是否与列名声明相符 ─────────────────────────
    for col in p.columns:
        if col.inferred_kind != "数值" or not col.stats:
            continue
        du = declared_unit(col.name)
        if not du:
            continue
        unit, scale = du
        med = col.stats.get("median", 0)
        if med <= 0:
            continue
        # 如果声明是「万元」，中位数却在百万量级，多半是单位写错或数据没换算
        implied = med * scale
        suspicious = False
        note = ""
        if unit != "元" and med > MAGNITUDE_FACTOR * 1e4:
            suspicious = True
            note = (f"列名声明单位是「{unit}」，但中位数已经是 {med:,.0f}，"
                    f"换算成元是 {implied:,.0f} —— 量级可疑，"
                    f"很可能数值本身就是「元」而列名没改")
        if suspicious:
            rep.add(Check(f"单位一致性 · {col.letter}列「{col.name}」",
                          False, note, severity="error",
                          evidence={"median": med, "declared_unit": unit}))

    # ── 检查 5：Twyman's Law —— 异常值必须被解释 ───────────────────
    for col in p.columns:
        s = col.stats
        if not s or not s.get("outlier_count"):
            continue
        ex = ", ".join(f"{x:,.0f}" for x in s["outlier_examples"])
        rep.add(Check(
            f"异常值 · {col.letter}列「{col.name}」",
            True,
            f"{s['outlier_count']} 个值落在 IQR 1.5 倍之外：{ex}。"
            f"Twyman's Law：看起来异常的数字通常是错的 —— "
            f"逐个确认它们是真实业务（大客户、大促）还是脏数据（单位错、多打一个零）。",
            severity="warn",
            evidence={"outliers": s["outlier_examples"]},
        ))

    # ── 检查 6：外部声称核对 ───────────────────────────────────────
    if claims:
        by_name = {c.name: c for c in p.columns}
        for key, claimed in claims.items():
            col = by_name.get(key)
            if col is None:
                # 模糊匹配
                cands = [c for c in p.columns if key in c.name]
                col = cands[0] if len(cands) == 1 else None
            if col is None:
                rep.add(Check(f"声称核对 · {key}", False,
                              f"找不到名为「{key}」的列，无法核对", severity="error"))
                continue
            ci = col.index
            actual = 0.0
            for r in detail_rows:
                v, _ = parse_number(grid[r - 1][ci]) if ci < len(grid[r - 1]) else (None, [])
                actual += v or 0.0
            rel = abs(actual - claimed) / abs(claimed) if claimed else float("inf")
            rep.add(Check(
                f"声称核对 · {key}",
                rel <= REL_TOL,
                f"你说 {claimed:,.2f}，明细求和 {actual:,.2f}"
                + ("  ✓" if rel <= REL_TOL else f"  ✗ 差 {actual - claimed:,.2f}（{rel:.2%}）"),
                evidence={"claimed": claimed, "actual": actual},
            ))

    return rep


# ── 输出 ──────────────────────────────────────────────────────────────
def render(rep: VerifyReport) -> str:
    L: list[str] = []
    add = L.append
    bar = "─" * 68
    add(bar)
    add(f"数字对账：{Path(rep.path).name}  [{rep.sheet}]")
    add(bar)

    add("")
    add("【口径声明】结论的前提，先说清楚")
    for k, v in rep.caliber.items():
        add(f"  {k}：{v}")

    errs = [c for c in rep.checks if not c.passed and c.severity == "error"]
    warns = [c for c in rep.checks if c.severity == "warn"]
    infos = [c for c in rep.checks if c.passed and c.severity == "info"]

    if errs:
        add("")
        add("【对不上 —— 必须解决】")
        for c in errs:
            add(f"  ✗ {c.name}")
            add(f"      {c.detail}")

    if infos:
        add("")
        add("【对上了】")
        for c in infos:
            add(f"  ✓ {c.name}")
            add(f"      {c.detail}")

    if warns:
        add("")
        add("【需要你做判断 —— 脚本判不了的部分】")
        for c in warns:
            add(f"  ? {c.name}")
            add(f"      {c.detail}")

    add("")
    add(bar)
    if rep.master_check:
        add("MASTER CHECK：通过。所有可机器验证的项目都对得上。")
        add("注意：机器只能验证「算得对不对」，验证不了「这个分析有没有意义」。")
    else:
        add(f"MASTER CHECK：不通过。{len(errs)} 项对不上，见上。")
        add("在解决之前，不要把这些数字写进任何交付物。")
    add(bar)
    return "\n".join(L)


def main() -> None:
    ap = argparse.ArgumentParser(description="数字对账：用表自己的数据证明你没算错")
    ap.add_argument("file")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--claim", action="append", default=[],
                    help='核对你自己算的数，格式 "列名:数值"，可重复')
    a = ap.parse_args()

    path = Path(a.file).expanduser()
    if not path.exists():
        sys.exit(f"文件不存在：{path}")

    claims: dict[str, float] = {}
    for c in a.claim:
        if ":" not in c and "：" not in c:
            sys.exit(f'--claim 格式应为 "列名:数值"，收到：{c}')
        k, _, v = c.replace("：", ":").partition(":")
        try:
            claims[k.strip()] = float(v.replace(",", "").strip())
        except ValueError:
            sys.exit(f"无法解析数值：{v}")

    rep = verify(path, a.sheet, claims or None)
    if a.json:
        print(json.dumps(asdict(rep), ensure_ascii=False, indent=2, default=str))
    else:
        print(render(rep))
    sys.exit(0 if rep.master_check else 1)


if __name__ == "__main__":
    main()
