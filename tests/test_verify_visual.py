"""verify_visual.py 的回归测试（需要 playwright + chromium，否则整体跳过）。

锁定本次修复的两个已知盲区：
  · markOverlap —— 图形之间部分重叠（原只报文字相关重叠，漏掉 mark-on-mark）
  · numbers.meta —— 采集 aria-label / <title> 里的数字（原完全看不见）

本地未装 playwright 时本文件会被 pytest 整体跳过；在 CI / 作者机器
（已装 playwright）上会自动跑，确保这两个盲区不再回退。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / "scripts"
sys.path.insert(0, str(SCRIPTS))

pytest.importorskip("playwright")  # 没装浏览器就跳过，不报错

import os  # noqa: E402

# 本地若只装了完整 chromium（未装 headless shell），自动指向其 chrome.exe，
# 让测试免等 headless shell 下载也能跑；CI/作者机器有 headless shell 时不设，
# 走默认启动（verify_visual 仅在 PW_CHROMIUM_EXE 非空时改用 executable_path）。
def _auto_exe() -> str:
    """挑版本号最大的完整 chromium（chrome.exe）作为 executable_path。

    完整 chromium 本身支持 headless，无需 headless shell。这样本地只要
    装了完整 chromium 就能跑测试，不必等 headless shell 下载；CI 上
    `playwright install chromium` 也会装完整 chromium，同样可用。
    """
    import re
    base = Path(os.environ.get("LOCALAPPDATA", "")) / "ms-playwright"
    cands = list(base.glob("chromium-*/chrome-win64/chrome.exe"))
    if not cands:
        return ""
    def ver(p: Path) -> int:
        m = re.search(r"chromium-(\d+)", str(p))
        return int(m.group(1)) if m else 0
    return str(max(cands, key=ver))

_exe = _auto_exe()
if _exe:
    os.environ["PW_CHROMIUM_EXE"] = _exe

from playwright.sync_api import sync_playwright  # noqa: E402
import verify_visual as vv  # noqa: E402


HTML = """<!doctype html><html><body>
<svg viewBox="0 0 100 100" aria-label="总营收 12340 万元">
  <rect x="10" y="10" width="30" height="30" fill="#0173B2"/>
  <rect x="25" y="10" width="30" height="30" fill="#DE8F05"/>
  <text x="5" y="60">12345</text>
</svg>
</body></html>"""


def _run_checks(html_path: Path) -> dict:
    with sync_playwright() as p:
        _exe = os.environ.get("PW_CHROMIUM_EXE") or None
        b = p.chromium.launch(executable_path=_exe)
        pg = b.new_page()
        pg.goto("file://" + str(html_path))
        pg.wait_for_timeout(300)
        out = pg.evaluate(vv.CHECKS)
        b.close()
    return out


def test_mark_overlap_detected(tmp_path):
    """两个矩形重叠 50% 必须被 markOverlap 抓到（盲区1）。"""
    html = tmp_path / "r.html"
    html.write_text(HTML, encoding="utf-8")
    out = _run_checks(html)
    assert out["markOverlap"], f"部分重叠未被检测：{out}"


def test_aria_label_numbers_visible(tmp_path):
    """aria-label 里的数字必须进入 numbers.meta（盲区2：原来看不见）。"""
    html = tmp_path / "r.html"
    html.write_text(HTML, encoding="utf-8")
    out = _run_checks(html)
    assert out["numbers"]["meta"], f"aria-label 数字未被采集：{out['numbers']}"
    assert 12340.0 in out["numbers"]["meta"], "aria-label 的 12340 应被采到"


def test_aria_label_mismatch_flagged(tmp_path):
    """aria-label 数字与图内数字接近但不等，必须被 nearMissMeta 标出。"""
    html = tmp_path / "r.html"
    html.write_text(HTML, encoding="utf-8")
    out = _run_checks(html)
    meta = out["numbers"]["meta"]
    body = out["numbers"]["svg"] + out["numbers"]["body"]
    hits = vv.near_miss(meta, body)
    assert hits, f"aria-label 12340 与图内 12345 的偏差未被标出：{hits}"
