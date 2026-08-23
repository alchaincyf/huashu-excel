#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把洞察渲染成一份自包含的 HTML 报告。

分工
----
洞察由 agent 产出（那是它擅长的），本脚本只负责渲染和画图——
这样图表的诚实性约束、口径声明、边界声明这几件容易被省略的事，
就变成了结构性保证而不是靠自觉。

零外部资源
----------
无 CDN、无外部字体、无网络请求。图表是自己画的内联 SVG。
所以它在 file:// 下、在离线环境、在内网都能正常打开，
也能直接当附件发给别人。

用法
----
    python3 make_report.py --spec report.json --out report.html
    python3 make_report.py --template > report.json      # 拿一份模板改

spec 结构见 --template。风格由 agent 按内容判断后填进 spec，
不要拿去问用户——报告不是设计作品，选风格是你的活。

依赖：无。纯标准库。
"""

from __future__ import annotations

import argparse
import html
import json
import sys
from pathlib import Path
from typing import Any

# ── 风格库 ──────────────────────────────────────────────────────────
# 每种风格取自一类机构的设计语言。真正的区别不在配色，在结构：
# 咨询给每张图编号并把结论写进图题，投行把表格做到最高密度，
# 财经媒体用衬线和窄栏做叙事，科技公司靠留白和圆角。
#
# 选哪种是 agent 的活，不是用户的活（见 references/report-styles.md）。
# 都不合适就用 custom 自己配。
SANS = "-apple-system,'PingFang SC','Microsoft YaHei','Hiragino Sans GB',Helvetica,sans-serif"
SERIF = "'Songti SC','Source Han Serif SC','Noto Serif CJK SC',Georgia,'Times New Roman',serif"

STYLES = {
    "consulting": {
        "desc": "咨询公司。深海军蓝，每张图编号 Exhibit N，图题即结论。战略分析、框架评估、董事会材料",
        "bg": "#FFFFFF", "fg": "#051C2C", "muted": "#5A6E7C",
        "rule": "#D6DEE4", "card": "#F4F7F9", "accent": "#2251FF",
        "pos": "#00806A", "neg": "#C1372B",
        "h_font": SANS, "b_font": SANS,
        "exhibit": True,
        "extra_css": """
h1{letter-spacing:-.02em;font-weight:700;border-bottom:3px solid %(accent)s;
padding-bottom:14px;display:inline-block}
h2{border-bottom:0;font-size:19px;color:%(fg)s;
padding-left:12px;border-left:4px solid %(accent)s}
.chart-title{font-size:16px;line-height:1.5}
.exhibit{font-size:11px;letter-spacing:.11em;text-transform:uppercase;
color:%(accent)s;font-weight:700;margin-bottom:5px}
.lede{border-left-width:4px;background:%(card)s}
.kpi{border-top:3px solid %(accent)s;border-radius:0}
th{text-transform:uppercase;letter-spacing:.05em;font-size:11.5px}
""",
    },
    "bank": {
        "desc": "投行研究报告。深蓝、高信息密度、等宽数字、细密分隔线。财务建模、估值、对账、尽调",
        "bg": "#FFFFFF", "fg": "#0B1F3A", "muted": "#63748A",
        "rule": "#D3DAE3", "card": "#F5F7FA", "accent": "#0B4F8C",
        "pos": "#0F6B3D", "neg": "#A01B22",
        "h_font": SANS, "b_font": SANS,
        "extra_css": """
body{font-size:15px;line-height:1.65}
.wrap{max-width:920px}
h1{font-size:26px;font-weight:700}
h2{font-size:17px;margin:38px 0 10px;text-transform:none;
border-bottom:2px solid %(fg)s;padding-bottom:6px}
table{font-size:13px}
th,td{padding:6px 10px}
tbody tr:nth-child(even){background:%(card)s}
.kpi{border-radius:3px;padding:12px 14px}
.kpi .v{font-size:22px}
.caliber{border-radius:3px;font-size:13px}
.lede{border-radius:0;font-size:15.5px}
""",
    },
    "editorial": {
        "desc": "财经媒体。三文鱼粉底、衬线标题、窄栏叙事。行业观察、深度解读、有观点的报告",
        "bg": "#FFF1E5", "fg": "#33302E", "muted": "#66605C",
        "rule": "#E6D9CB", "card": "#FFFAF5", "accent": "#0F5499",
        "pos": "#0D7680", "neg": "#CC0000",
        "h_font": SERIF, "b_font": SANS,
        "extra_css": """
.wrap{max-width:760px}
h1{font-size:34px;line-height:1.28}
h2{font-family:%(h_font)s;border-bottom:0;font-size:22px;
margin:48px 0 12px}
p{font-size:16.5px;line-height:1.8}
.lede{background:transparent;border-left:3px solid %(accent)s;
font-family:%(h_font)s;font-size:18px}
.kpi{background:%(card)s;border-color:%(rule)s}
.chart{border-top:1px solid %(rule)s;padding-top:16px}
""",
    },
    "magazine": {
        "desc": "杂志式。白底红标块、紧凑排版、editorial 标题。观点报告、行业洞察、有立场的分析",
        "bg": "#FFFFFF", "fg": "#121212", "muted": "#6E6E6E",
        "rule": "#E0E0E0", "card": "#F6F6F6", "accent": "#E3120B",
        "pos": "#0B7A3B", "neg": "#E3120B",
        "h_font": SANS, "b_font": SANS,
        "extra_css": """
.wrap{max-width:740px}
h1{font-size:29px;line-height:1.3;font-weight:700}
h1::before{content:"";display:block;width:52px;height:7px;
background:%(accent)s;margin-bottom:15px}
h2{border-bottom:0;font-size:18px;margin:44px 0 10px;
padding-left:0}
h2::before{content:"";display:inline-block;width:14px;height:3px;
background:%(accent)s;vertical-align:middle;margin-right:9px}
.chart-title::before{content:"";display:inline-block;width:9px;height:9px;
background:%(accent)s;margin-right:8px}
body{font-size:15.5px}
.lede{background:%(card)s;border-left:3px solid %(accent)s}
""",
    },
    "product": {
        "desc": "科技公司。大留白、圆角卡片、克制的紫蓝强调。产品复盘、增长分析、内部评审",
        "bg": "#FFFFFF", "fg": "#0A2540", "muted": "#6B7C93",
        "rule": "#E6EBF1", "card": "#F6F9FC", "accent": "#635BFF",
        "pos": "#09825D", "neg": "#CD3D64",
        "h_font": SANS, "b_font": SANS,
        "extra_css": """
.wrap{max-width:840px;padding-top:72px}
h1{font-size:33px;letter-spacing:-.025em;font-weight:700}
h2{border-bottom:0;font-size:21px;margin:56px 0 12px;letter-spacing:-.015em}
.kpi{border-radius:12px;border-color:%(rule)s;background:%(card)s;padding:18px 20px}
.kpi .v{letter-spacing:-.03em}
.lede{border-radius:12px;border-left:0;background:%(card)s;padding:22px 24px}
.caliber{border-radius:12px}
.bounds{border-radius:12px}
.chart{background:%(card)s;border-radius:12px;padding:20px 22px}
table{font-size:14.5px}
""",
    },
    "minimal": {
        "desc": "极简。大量留白、无装饰。内部快看、单一结论、只想把数说清楚",
        "bg": "#FFFFFF", "fg": "#000000", "muted": "#767676",
        "rule": "#EAEAEA", "card": "#FAFAFA", "accent": "#000000",
        "pos": "#1F6F3F", "neg": "#95291F",
        "h_font": SANS, "b_font": SANS,
        "extra_css": """
.wrap{max-width:720px}
h2{border-bottom:0;font-size:18px;margin:48px 0 10px}
.lede{background:transparent;border-left:2px solid %(fg)s}
.kpi{border:0;background:transparent;padding:0 20px 0 0}
.kpis{gap:32px}
""",
    },
}
# 别名，向后兼容
STYLES["report"] = STYLES["editorial"]
STYLES["finance"] = STYLES["bank"]

# 色盲友好色板（Wong 系）：避开红绿对立，约 8% 男性有红绿色觉障碍
SERIES_COLORS = ["#0173B2", "#DE8F05", "#029E73", "#CC78BC", "#CA9161", "#56B4E9"]

E = html.escape


def _nice_num(x: float, round_it: bool) -> float:
    """把一个数收成「好看的数」：1 / 2 / 2.5 / 5 的 10 的幂倍。"""
    import math
    if x <= 0:
        return 1.0
    exp = math.floor(math.log10(x))
    f = x / (10 ** exp)
    if round_it:
        nf = 1 if f < 1.5 else (2 if f < 3 else (5 if f < 7 else 10))
    else:
        nf = 1 if f <= 1 else (2 if f <= 2 else (5 if f <= 5 else 10))
    return nf * (10 ** exp)


def nice_scale(lo: float, hi: float, ticks: int = 4) -> tuple[float, float, float]:
    """算出整齐的坐标轴刻度。

    139.68 / 115.50 / 91.32 这种从数据直接算出来的碎刻度，
    读者要多花一秒去解析每个数字。整刻度（100/120/140）是免费的可读性。

    注意：这里只把刻度收整，不强制从 0 —— 折线编码的是位置和斜率，
    强行归零会把真实波动压平，那是另一种误导。
    """
    import math
    if hi <= lo:
        hi, lo = lo + 1, lo - 1
    span = _nice_num(hi - lo, False)
    step = _nice_num(span / max(ticks - 1, 1), True)
    return math.floor(lo / step) * step, math.ceil(hi / step) * step, step


def _fmt(v: float) -> str:
    if v is None:
        return "—"
    a = abs(v)
    if a >= 1e8:
        return f"{v/1e8:,.2f}亿"
    if a >= 1e4:
        return f"{v/1e4:,.1f}万"
    if a == int(a):
        return f"{int(v):,}"
    return f"{v:,.2f}"


# ── SVG 图表：自己画，不依赖任何库 ──────────────────────────────────
def svg_bar(labels: list[str], values: list[float], st: dict,
            highlight: int | None = None) -> str:
    """横向条形图。长度 + 共同基线是感知精度最高的编码方式
    （Cleveland & McGill 1984），所以这是默认图表。"""
    if not values:
        return ""
    n = len(values)
    row_h, gap, pad_l, pad_r, pad_t = 30, 8, 130, 90, 8
    h = pad_t * 2 + n * row_h + (n - 1) * gap
    w = 720
    plot_w = w - pad_l - pad_r
    # 基线从 0：条形编码长度，截断即失真
    vmax = max(max(values), 0)
    vmin = min(min(values), 0)
    span = (vmax - vmin) or 1
    zero_x = pad_l + (0 - vmin) / span * plot_w

    parts = [f'<svg viewBox="0 0 {w} {h}" width="100%" '
             f'preserveAspectRatio="xMidYMid meet" role="img">']
    for i, (lab, val) in enumerate(zip(labels, values)):
        y = pad_t + i * (row_h + gap)
        x1 = pad_l + (min(val, 0) - vmin) / span * plot_w
        x2 = pad_l + (max(val, 0) - vmin) / span * plot_w
        bw = max(x2 - x1, 1.5)
        color = st["accent"] if (highlight is not None and i == highlight) \
            else SERIES_COLORS[0]
        op = "1" if (highlight is None or i == highlight) else "0.45"
        parts.append(
            f'<rect x="{x1:.1f}" y="{y}" width="{bw:.1f}" height="{row_h}" '
            f'rx="2" fill="{color}" opacity="{op}"/>')
        parts.append(
            f'<text x="{pad_l - 10}" y="{y + row_h*0.68:.0f}" text-anchor="end" '
            f'font-size="13" fill="{st["fg"]}">{E(str(lab)[:12])}</text>')
        # 直接标注数值，不用图例也不用坐标轴刻度
        parts.append(
            f'<text x="{x2 + 8:.1f}" y="{y + row_h*0.68:.0f}" font-size="13" '
            f'font-variant-numeric="tabular-nums" fill="{st["muted"]}">{_fmt(val)}</text>')
    if vmin < 0:
        parts.append(f'<line x1="{zero_x:.1f}" y1="{pad_t-4}" x2="{zero_x:.1f}" '
                     f'y2="{h-pad_t+4}" stroke="{st["rule"]}" stroke-width="1"/>')
    parts.append("</svg>")
    return "".join(parts)


def svg_line(labels: list[str], series: list[dict], st: dict,
             incomplete_last: bool = False) -> str:
    """折线图。不强制 Y 轴从 0——折线编码位置和斜率，强行归零会压平真实波动。"""
    if not series or not series[0].get("values"):
        return ""
    w, h = 720, 300
    pad_l, pad_r, pad_t, pad_b = 56, 80, 16, 34
    pw, ph = w - pad_l - pad_r, h - pad_t - pad_b
    allv = [v for s in series for v in s["values"] if v is not None]
    if not allv:
        return ""
    dlo, dhi = min(allv), max(allv)
    pad = (dhi - dlo) * 0.12 if dhi > dlo else max(abs(dhi) * 0.1, 1)
    lo, hi, step = nice_scale(dlo - pad, dhi + pad, ticks=4)
    n = len(labels)

    def X(i: int) -> float:
        return pad_l + (i / max(n - 1, 1)) * pw

    def Y(v: float) -> float:
        return pad_t + (1 - (v - lo) / (hi - lo)) * ph

    parts = [f'<svg viewBox="0 0 {w} {h}" width="100%" '
             f'preserveAspectRatio="xMidYMid meet" role="img">']
    # 极淡的水平参考线，落在整刻度上
    val = lo
    while val <= hi + step * 0.001:
        y = Y(val)
        parts.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{pad_l+pw}" y2="{y:.1f}" '
                     f'stroke="{st["rule"]}" stroke-width="1"/>')
        parts.append(f'<text x="{pad_l-8}" y="{y+4:.1f}" text-anchor="end" '
                     f'font-size="11" fill="{st["muted"]}" '
                     f'font-variant-numeric="tabular-nums">{_fmt(val)}</text>')
        val += step

    for si, s in enumerate(series):
        color = SERIES_COLORS[si % len(SERIES_COLORS)]
        vals = s["values"]
        pts = [(X(i), Y(v)) for i, v in enumerate(vals) if v is not None]
        if not pts:
            continue
        # 不完整的最后一期用虚线，而不是画成一次真实的下跌
        if incomplete_last and len(pts) > 1:
            solid = pts[:-1]
            parts.append('<polyline points="'
                         + " ".join(f"{x:.1f},{y:.1f}" for x, y in solid)
                         + f'" fill="none" stroke="{color}" stroke-width="2.2" '
                           f'stroke-linejoin="round"/>')
            (x1, y1), (x2, y2) = pts[-2], pts[-1]
            parts.append(f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
                         f'stroke="{color}" stroke-width="2.2" stroke-dasharray="5 4"/>')
        else:
            parts.append('<polyline points="'
                         + " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
                         + f'" fill="none" stroke="{color}" stroke-width="2.2" '
                           f'stroke-linejoin="round"/>')
        for x, y in pts:
            parts.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="{color}"/>')
        # 直接在线末标注系列名，不做图例
        lx, ly = pts[-1]
        parts.append(f'<text x="{lx+9:.1f}" y="{ly+4:.1f}" font-size="12" '
                     f'fill="{color}">{E(str(s.get("name", "")))[:10]}</text>')

    step = max(1, n // 8)
    for i, lab in enumerate(labels):
        if i % step == 0 or i == n - 1:
            parts.append(f'<text x="{X(i):.1f}" y="{h-12}" text-anchor="middle" '
                         f'font-size="11" fill="{st["muted"]}">{E(str(lab)[:10])}</text>')
    parts.append("</svg>")
    return "".join(parts)


def render_chart(ch: dict, st: dict, exhibit_no: int | None = None) -> str:
    kind = ch.get("type", "bar")
    labels = [str(x) for x in ch.get("labels", [])]
    if kind == "line":
        series = ch.get("series") or [{"name": ch.get("name", ""),
                                       "values": ch.get("values", [])}]
        body = svg_line(labels, series, st, ch.get("incomplete_last", False))
    else:
        vals = ch.get("values", [])
        pairs = list(zip(labels, vals))
        if ch.get("sort", True):
            pairs.sort(key=lambda t: (t[1] is None, -(t[1] or 0)))
        hl = ch.get("highlight")
        hi = None
        if hl is not None:
            for i, (lab, _v) in enumerate(pairs):
                if lab == str(hl):
                    hi = i
                    break
        body = svg_bar([p[0] for p in pairs], [p[1] for p in pairs], st, hi)
    if not body:
        return ""
    cap = ch.get("caption", "")
    note = ch.get("note", "")
    out = ['<figure class="chart">']
    if exhibit_no is not None:
        out.append(f'<div class="exhibit">Exhibit {exhibit_no}</div>')
    if cap:
        out.append(f'<figcaption class="chart-title">{E(cap)}</figcaption>')
    out.append(f'<div class="chart-body">{body}</div>')
    if note:
        out.append(f'<p class="chart-note">{E(note)}</p>')
    out.append("</figure>")
    return "".join(out)


def render_table(tb: dict, st: dict) -> str:
    cols = tb.get("columns", [])
    rows = tb.get("rows", [])
    if not cols or not rows:
        return ""
    out = ['<div class="table-wrap"><table>']
    if tb.get("caption"):
        out.append(f'<caption>{E(tb["caption"])}</caption>')
    out.append("<thead><tr>" + "".join(f"<th>{E(str(c))}</th>" for c in cols)
               + "</tr></thead><tbody>")
    for r in rows:
        cells = []
        for v in r:
            num = isinstance(v, (int, float)) and not isinstance(v, bool)
            cells.append(f'<td class="{"num" if num else ""}">'
                         f'{_fmt(v) if num else E(str(v))}</td>')
        out.append("<tr>" + "".join(cells) + "</tr>")
    out.append("</tbody></table></div>")
    return "".join(out)


CSS = """
*{box-sizing:border-box}
body{margin:0;background:%(bg)s;color:%(fg)s;font-family:%(b_font)s;
line-height:1.75;font-size:16px;-webkit-font-smoothing:antialiased}
.wrap{max-width:860px;margin:0 auto;padding:56px 28px 96px}
h1{font-family:%(h_font)s;font-size:31px;line-height:1.35;margin:0 0 10px;
font-weight:700;letter-spacing:-.01em}
h2{font-family:%(h_font)s;font-size:20px;margin:52px 0 14px;font-weight:700;
padding-bottom:9px;border-bottom:1px solid %(rule)s}
h3{font-size:16px;margin:26px 0 8px;font-weight:600}
p{margin:0 0 15px}
.sub{color:%(muted)s;font-size:14px;margin-bottom:30px}
.lede{font-size:17px;padding:20px 22px;background:%(card)s;
border-left:3px solid %(accent)s;border-radius:0 6px 6px 0;margin:0 0 32px}
.lede ul{margin:0;padding-left:19px}
.lede li{margin:7px 0}
.kpis{display:flex;flex-wrap:wrap;gap:12px;margin:0 0 30px}
.kpi{flex:1 1 150px;background:%(card)s;border:1px solid %(rule)s;
border-radius:8px;padding:15px 17px}
.kpi .k{font-size:12px;color:%(muted)s;margin-bottom:5px}
.kpi .v{font-size:25px;font-weight:700;font-variant-numeric:tabular-nums;
letter-spacing:-.02em}
.kpi .d{font-size:12px;color:%(muted)s;margin-top:3px}
.pos{color:%(pos)s}.neg{color:%(neg)s}
.chart{margin:22px 0 26px}
.chart-title{font-size:15px;font-weight:600;margin-bottom:10px}
.chart-body{overflow-x:auto}
.chart-note{font-size:12.5px;color:%(muted)s;margin:8px 0 0}
.table-wrap{overflow-x:auto;margin:18px 0 24px}
table{border-collapse:collapse;width:100%%;font-size:14px}
caption{text-align:left;font-size:13px;color:%(muted)s;padding-bottom:8px}
th,td{padding:8px 11px;border-bottom:1px solid %(rule)s;text-align:left}
th{font-weight:600;font-size:13px;color:%(muted)s;
border-bottom:1.5px solid %(rule)s}
td.num{text-align:right;font-variant-numeric:tabular-nums}
tbody tr:hover{background:%(card)s}
.caliber{background:%(card)s;border:1px solid %(rule)s;border-radius:8px;
padding:17px 20px;margin:0 0 30px;font-size:14px}
.caliber dl{margin:0;display:grid;grid-template-columns:auto 1fr;gap:5px 16px}
.caliber dt{color:%(muted)s;white-space:nowrap}
.caliber dd{margin:0}
.bounds{border:1px dashed %(rule)s;border-radius:8px;padding:17px 20px;
margin:40px 0 0;font-size:14px;color:%(muted)s}
.bounds h2{border:0;margin:0 0 9px;font-size:15px;padding:0;color:%(fg)s}
.bounds ul{margin:0;padding-left:19px}
.bounds li{margin:5px 0}
footer{margin-top:44px;padding-top:16px;border-top:1px solid %(rule)s;
font-size:12px;color:%(muted)s}
@media(max-width:600px){.wrap{padding:32px 18px 64px}h1{font-size:25px}}
@media print{body{background:#fff}.wrap{padding:0}
.chart,.table-wrap,.kpi{break-inside:avoid}}
"""


def resolve_style(spec: dict) -> dict:
    """定出这份报告用什么风格。

    优先级：spec.custom_style（agent 自己配的）> spec.style（内置名）> consulting。
    内置的都不合适时，agent 应该直接给 custom_style，不要将就一个不对的。
    """
    base = dict(STYLES.get(spec.get("style", "consulting"), STYLES["consulting"]))
    custom = spec.get("custom_style")
    if isinstance(custom, dict):
        # 只允许覆盖表现层字段，防止 spec 注入结构性内容
        allowed = {"bg", "fg", "muted", "rule", "card", "accent", "pos", "neg",
                   "h_font", "b_font", "extra_css", "exhibit", "desc"}
        base.update({k: v for k, v in custom.items() if k in allowed})
    base.setdefault("extra_css", "")
    base.setdefault("exhibit", False)
    return base


def build_html(spec: dict) -> str:
    st = resolve_style(spec)
    exhibit_on = bool(st.get("exhibit")) or bool(spec.get("exhibit"))
    exhibit_n = 0

    body: list[str] = []
    body.append(f'<h1>{E(spec.get("title", "数据分析报告"))}</h1>')
    if spec.get("subtitle"):
        body.append(f'<p class="sub">{E(spec["subtitle"])}</p>')

    kpis = spec.get("kpis") or []
    if kpis:
        cards = []
        for k in kpis:
            val = k.get("value")
            vs = _fmt(val) if isinstance(val, (int, float)) else E(str(val))
            cls = ""
            d = k.get("delta")
            if isinstance(d, (int, float)):
                cls = "pos" if d > 0 else ("neg" if d < 0 else "")
            dtxt = ""
            if k.get("note"):
                dtxt = f'<div class="d {cls}">{E(str(k["note"]))}</div>'
            cards.append(f'<div class="kpi"><div class="k">{E(str(k.get("label","")))}'
                         f'</div><div class="v">{vs}</div>{dtxt}</div>')
        body.append('<div class="kpis">' + "".join(cards) + "</div>")

    summary = spec.get("summary") or []
    if summary:
        body.append('<div class="lede"><ul>'
                    + "".join(f"<li>{E(s)}</li>" for s in summary) + "</ul></div>")

    # 口径声明。结构性地放在正文之前——说不清口径的数字不能被检验。
    cal = spec.get("caliber") or {}
    if cal:
        items = "".join(f"<dt>{E(str(k))}</dt><dd>{E(str(v))}</dd>"
                        for k, v in cal.items())
        body.append(f'<div class="caliber"><dl>{items}</dl></div>')

    for sec in spec.get("sections", []):
        if sec.get("heading"):
            body.append(f'<h2>{E(sec["heading"])}</h2>')
        for para in (sec.get("body") or "").split("\n\n"):
            if para.strip():
                body.append(f"<p>{E(para.strip())}</p>")
        for ch in (sec.get("charts") or ([sec["chart"]] if sec.get("chart") else [])):
            if exhibit_on:
                exhibit_n += 1
            body.append(render_chart(ch, st, exhibit_n if exhibit_on else None))
        for tb in (sec.get("tables") or ([sec["table"]] if sec.get("table") else [])):
            body.append(render_table(tb, st))
        for sub in sec.get("subsections", []):
            if sub.get("heading"):
                body.append(f'<h3>{E(sub["heading"])}</h3>')
            for para in (sub.get("body") or "").split("\n\n"):
                if para.strip():
                    body.append(f"<p>{E(para.strip())}</p>")

    # 边界声明。固定结构，不由 agent 决定要不要写。
    bounds = spec.get("boundaries") or []
    if bounds:
        body.append('<div class="bounds"><h2>这份分析的边界</h2><ul>'
                    + "".join(f"<li>{E(b)}</li>" for b in bounds) + "</ul></div>")

    if spec.get("footer"):
        body.append(f'<footer>{E(spec["footer"])}</footer>')

    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{E(spec.get('title', '数据分析报告'))}</title>
<style>{CSS % st}{(st.get('extra_css') or '') % st}</style></head>
<body><div class="wrap">{''.join(body)}</div></body></html>"""


TEMPLATE = {
    "title": "标题写结论，不写主题（「华东贡献六成销售额但增速垫底」而不是「销售额分析」）",
    "subtitle": "数据范围 / 制表日期",
    "style": "report",
    "_style_note": "report=默认业务报告 | finance=财务审计类 | minimal=内部快看。"
                   "你按内容自己选，不要拿去问用户。",
    "kpis": [
        {"label": "上半年销售额", "value": 14688217, "note": "较去年同期 +12.4%", "delta": 1}
    ],
    "summary": [
        "第一条结论（能改变某个决定的那种，不是描述）",
        "第二条结论"
    ],
    "caliber": {
        "时间窗口": "2026-01-01 至 2026-06-30，按支付时间",
        "口径": "含税，不含运费；已排除汇总行；未去重",
        "数据来源": "销售明细.xlsx"
    },
    "sections": [
        {
            "heading": "小标题也写结论",
            "body": "段落。空行分段。\n\n第二段。",
            "chart": {
                "type": "bar",
                "caption": "图表标题写结论",
                "labels": ["华东", "华南", "华北"],
                "values": [6497500, 4556300, 2654317],
                "sort": True,
                "highlight": "华东",
                "note": "可选的图下注解：只写图的意义，不复述图里已有的数字"
            }
        },
        {
            "heading": "趋势",
            "body": "",
            "chart": {
                "type": "line",
                "caption": "三月起连续下滑",
                "labels": ["1月", "2月", "3月", "4月", "5月", "6月"],
                "series": [{"name": "销售额", "values": [120, 135, 128, 119, 108, 96]}],
                "incomplete_last": False,
                "note": "incomplete_last=true 会把最后一段画成虚线——"
                        "本期没走完却画成实线，是最常见的误导"
            },
            "table": {
                "caption": "可选表格",
                "columns": ["门店", "销售额", "客户数"],
                "rows": [["朝阳店", 1234567, 1203]]
            }
        }
    ],
    "boundaries": [
        "数据只覆盖 X 到 Y，结论在此之外不成立",
        "哪个结论最脆弱、换什么口径会翻盘",
        "哪些数字没有独立校验来源"
    ],
    "footer": "生成于 … ｜ 口径见上"
}


def main() -> None:
    ap = argparse.ArgumentParser(description="把洞察渲染成自包含 HTML 报告")
    ap.add_argument("--spec", help="洞察 spec 的 JSON 文件路径")
    ap.add_argument("--out", default="report.html")
    ap.add_argument("--template", action="store_true", help="打印一份 spec 模板")
    ap.add_argument("--styles", action="store_true", help="列出可用风格")
    a = ap.parse_args()

    if a.styles:
        for k, v in STYLES.items():
            print(f"{k:10} {v['desc']}")
        print("\n按内容自己选，不要拿去问用户——报告不是设计作品。")
        return
    if a.template:
        print(json.dumps(TEMPLATE, ensure_ascii=False, indent=2))
        return
    if not a.spec:
        sys.exit("需要 --spec <json文件>，或用 --template 拿一份模板")

    sp = Path(a.spec).expanduser()
    if not sp.exists():
        sys.exit(f"文件不存在：{sp}")
    try:
        spec = json.loads(sp.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        sys.exit(f"spec 不是合法 JSON：{e}")

    out = Path(a.out).expanduser()
    out.write_text(build_html(spec), encoding="utf-8")
    size = out.stat().st_size / 1024

    print(f"已生成：{out}  ({size:.1f} KB，自包含，无外部资源)")
    missing = []
    if not spec.get("caliber"):
        missing.append("caliber（口径声明）—— 没有口径的数字无法被检验")
    if not spec.get("boundaries"):
        missing.append("boundaries（分析边界）—— 说出自身弱点的报告才可信")
    if not spec.get("summary"):
        missing.append("summary（结论先行）—— 读者不该自己去找结论")
    for m in missing:
        print(f"  ⚠ 缺 {m}")


if __name__ == "__main__":
    main()
