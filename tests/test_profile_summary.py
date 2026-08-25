"""回归测试：汇总/小计行检测（修复「只看首格」漏判）。

复现 huashu-excel 的已知 bug——中文报表里小计标签常在后续列
（A 列"华东" + B 列"华东小计"），原 profile_table 只扫每行首格，
导致整行被当明细吞掉、求和翻倍且悄无声息。

本测试固化一张合成脏表，断言：
1. profile 把「华东小计」「总计」标进 summary_rows，且不污染明细行数；
2. verify_numbers 据此把小计当免费校验和，不再产出误导性的「与总计差 61.85%」。
"""
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / "scripts"
sys.path.insert(0, str(SCRIPTS))
import profile_table as pt  # noqa: E402


def _build_dirty(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "销售明细"
    ws.append(["华东大区2026年销售汇总表"])
    ws.append(["制表：张三"])
    ws.append(["地区", "门店", "一季度", "", "", "退货额(元)"])
    ws.append(["", "", "1月", "2月", "3月", "退货额(元)"])
    rows = [
        ["华东", "门店A", 1000000, 1000000, 1000000, "—"],
        ["华东", "门店B", 1500000, 1500000, 1500000, "无"],
        ["华东", "门店C", 2000000, 2000000, 2000000, "N/A"],
        ["华东", "门店D", 1997500, 1990420, 2000000, 5000],
        ["华东", "门店A", 1000000, 1000000, 1000000, "—"],  # 粘贴重复
        ["华北", "门店E", 800000, 800000, 800000, ""],
        ["华北", "门店F", 900000, 900000, 900000, ""],
    ]
    for r in rows:
        ws.append(r)
    # 小计标签在 B 列，A 列填分组名「华东」——正是原 bug 的触发布局
    ws.append(["华东", "华东小计", 6497500, 6490410, 6500000, ""])
    ws.append(["", "总计", 9697500, 9690410, 9700000, ""])
    ws.append(["", "占比", 0.33, 0.33, 0.33, ""])
    wb.save(path)


def test_profile_detects_summary_rows_anywhere(tmp_path):
    path = tmp_path / "dirty_sales.xlsx"
    _build_dirty(path)

    prof = pt.profile(path)
    labels = {s["row"]: s["label"] for s in prof.summary_rows}

    subtotal_row = next((rn for rn, lb in labels.items() if "小计" in lb), None)
    total_row = next((rn for rn, lb in labels.items() if "总计" in lb), None)

    assert subtotal_row is not None, f"华东小计行未识别为汇总行，summary_rows={prof.summary_rows}"
    assert total_row is not None, f"总计行未识别为汇总行，summary_rows={prof.summary_rows}"
    # 7 条明细不应被小计/总计行污染
    assert prof.n_data_rows == 7, f"明细行数应为 7，实际 {prof.n_data_rows}"


def test_profile_formula_subtotal_detected(tmp_path):
    """公式信号层：无文字标签、纯 =SUM() 的表中段小计行也能被精确识别。

    用 4 列表避免触发「稀疏行」启发式（窄表下该启发式本就激进，与本次修复无关）。
    放在表中段（而非末尾）：末尾的无缓存公式小计行会被 read_grid 的尾部空行裁剪
    提前丢弃，那是 read_grid 的既有行为，不在本次修复范围。
    """
    path = tmp_path / "formula_subtotal.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); ws = wb.active
    ws.append(["区域", "产品", "销量", "金额"])
    ws.append(["华东", "A", 10, 100])
    ws.append(["华东", "B", 20, 200])
    ws.append(["", "", "", "=SUM(D2:D3)"])  # 无文字标签，表中段公式小计
    ws.append(["华北", "C", 30, 300])
    wb.save(path)
    prof = pt.profile(path)
    assert any("公式" in s.get("reason", "") for s in prof.summary_rows), \
        f"公式小计行未被识别，summary_rows={prof.summary_rows}"
    assert prof.n_data_rows == 3, f"明细行数应为 3，实际 {prof.n_data_rows}"


def test_verify_treats_subtotal_as_checksum(tmp_path):
    path = tmp_path / "dirty_sales.xlsx"
    _build_dirty(path)

    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "verify_numbers.py"), str(path)],
        capture_output=True, text=True, cwd=str(SCRIPTS),
    )
    out = proc.stdout + proc.stderr

    # 修复后：把华东小计当分组校验和，且不再产出旧的误导性「与总计差 61.85%」
    assert proc.returncode == 1, f"应检测到真实差额并退出码1，out={out[:500]}"
    assert "华东小计" in out, f"修复后应对华东小计做分组校验，out={out[:500]}"
    assert "61.85" not in out, f"不应再出现旧的错误'与总计差61.85%'，out={out[:500]}"
