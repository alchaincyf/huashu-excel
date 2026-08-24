#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 文档交付前的渲染自检——docx 版的闸门。

为什么需要它
------------
HTML 报告有 verify_visual.py 把关，docx 长期没有。于是这几类错
只能靠人肉发现，而它们**在生成端全都不报错**：

  · 中文 run 没设 w:eastAsia → Word 用主题字体兜底，整篇变宋体
  · 字体名是 Mac 或 Windows 独有 → 换台机器打开就回退
  · 「空段落 + 分页符」的写法 → 上一页正好排满时留下一整页全白
  · 表格跨页但表头没设重复 → 后续页面的表格没有列名
  · 图片宽过版心 → 右边被裁掉

检查项（机器可判的事实，不判好看不好看）
  FAIL  中文字符所在 run 缺 w:eastAsia
  FAIL  图片宽度超出版心
  FAIL  正文出现近乎空白的页（需要 soffice；缺了就跳过并说明）
  WARN  使用了跨平台风险字体
  WARN  空段落里放分页符（空白页的成因）
  WARN  表格跨页但未设置表头重复
  WARN  文档没有任何图片（图表被降级成表格时容易发生）

用法
    python3 verify_docx.py 报告.docx
    python3 verify_docx.py 报告.docx --shots 自检/p    额外导出每页 PNG
    python3 verify_docx.py 报告.docx --json

退出码：有 FAIL 返回 1，否则 0。可直接当流水线闸门。

依赖：解析部分纯标准库；渲染部分需要 LibreOffice（soffice）与 pdftoppm，
装不上会跳过渲染类检查并明确告诉你哪几项没验，不会假装通过。
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

CJK = re.compile(r"[㐀-鿿豈-﫿　-〿＀-￯]")

# Microsoft Office 在 Windows 与 macOS 上都随装的中日韩字体。
# 只有这一组能保证「换台机器打开还是这个样子」。
SAFE_EASTASIA = {
    "Microsoft YaHei", "微软雅黑", "Microsoft YaHei UI",
    "DengXian", "等线", "DengXian Light", "等线 Light",
    "SimSun", "宋体", "NSimSun", "新宋体", "SimHei", "黑体",
    "KaiTi", "楷体", "FangSong", "仿宋",
    "Microsoft JhengHei", "微軟正黑體",
    "MS Gothic", "MS Mincho", "Yu Gothic", "Meiryo",
    "Malgun Gothic",
}
# 常见的「只在一个平台上有」的字体，命中就提示
PLATFORM_ONLY = {
    "PingFang SC": "仅 macOS", "PingFang TC": "仅 macOS",
    "Hiragino Sans GB": "仅 macOS", "STHeiti": "仅 macOS",
    "Heiti SC": "仅 macOS", "Songti SC": "仅 macOS",
    "Helvetica": "仅 macOS（Windows 会回退到 Arial）",
    "Helvetica Neue": "仅 macOS",
    "Segoe UI": "仅 Windows", "Calibri": "Windows 与新版 Office",
    "Source Han Sans": "需自行安装", "Noto Sans CJK SC": "需自行安装",
}

EMU_PER_CM = 360000


def xml_health(path: Path) -> list[str]:
    """包结构合法性。后处理脚本（补图、改字体）最容易把包改坏——
    实测有人往 [Content_Types].xml 插节点插到了根元素外面，Word 和
    LibreOffice 都拒绝打开，而当时的闸门照样解析成功、全绿放行。
    所以先验包，包坏了后面的一切检查都没有意义。"""
    probs: list[str] = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        for req in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
            if req not in names:
                probs.append(f"缺关键部件 {req}，Word 无法打开")
        for n in sorted(names):
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    ET.fromstring(z.read(n))
                except ET.ParseError as e:
                    probs.append(f"{n} 不是合法 XML（{e}）——Word/LibreOffice 会拒绝打开，"
                                 "多半是后处理脚本把节点插到了根元素外面")
        # 图片引用三方一致：document 里的 r:embed ↔ rels ↔ media 实体文件
        if "word/_rels/document.xml.rels" in names and "word/document.xml" in names:
            try:
                relx = ET.fromstring(z.read("word/_rels/document.xml.rels"))
                rels = {r.get("Id"): r.get("Target") for r in relx}
                doc_raw = z.read("word/document.xml").decode("utf-8", "ignore")
                for rid in set(re.findall(r'r:embed="([^"]+)"', doc_raw)):
                    tgt = rels.get(rid)
                    if not tgt:
                        probs.append(f"图片引用 {rid} 在 rels 里没有对应关系，该图会显示为红叉")
                    elif f"word/{tgt.lstrip('/')}" not in names and tgt.lstrip("/") not in names:
                        probs.append(f"图片 {rid} 指向 {tgt}，包内没有这个文件")
            except ET.ParseError:
                pass    # 上面已经报过
    return probs


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def parse(path: Path) -> dict:
    """把 docx 拆出检查需要的事实。纯标准库。"""
    with zipfile.ZipFile(path) as z:
        doc = ET.fromstring(z.read("word/document.xml"))
        names = set(z.namelist())
        media = [n for n in names if n.startswith("word/media/")]
        styles = None
        if "word/styles.xml" in names:
            styles = ET.fromstring(z.read("word/styles.xml"))

    # 版心宽度（EMU）
    sect = doc.find(f".//{W}sectPr")
    text_w_emu = None
    if sect is not None:
        pg, mar = sect.find(f"{W}pgSz"), sect.find(f"{W}pgMar")
        if pg is not None and mar is not None:
            tw = int(pg.get(f"{W}w", 0))
            lm = int(mar.get(f"{W}left", 0)); rm = int(mar.get(f"{W}right", 0))
            text_w_emu = (tw - lm - rm) * 635          # twip → EMU

    # 样式继承：run 没写 rFonts 时会继承段落样式、再继承 docDefaults。
    # 只看 run 级别会把「字体声明在样式里」这种完全合法的写法误判为缺字体。
    style_ea: dict[str, str] = {}
    style_based: dict[str, str] = {}
    default_ea = None
    if styles is not None:
        dd = styles.find(f"{W}docDefaults")
        if dd is not None:
            for rf in dd.iter(f"{W}rFonts"):
                default_ea = rf.get(f"{W}eastAsia") or default_ea
        for st in styles.iter(f"{W}style"):
            sid = st.get(f"{W}styleId")
            if not sid:
                continue
            bo = st.find(f"{W}basedOn")
            if bo is not None and bo.get(f"{W}val"):
                style_based[sid] = bo.get(f"{W}val")
            rf = st.find(f"{W}rPr/{W}rFonts")
            if rf is not None and rf.get(f"{W}eastAsia"):
                style_ea[sid] = rf.get(f"{W}eastAsia")

    def resolve_style_ea(sid: str | None) -> str | None:
        """顺着 basedOn 链找 eastAsia，找不到就回落到 docDefaults。"""
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            if sid in style_ea:
                return style_ea[sid]
            sid = style_based.get(sid)
        return default_ea

    bad_runs, fonts, ea_fonts, empty_break = [], set(), set(), 0
    for p in doc.iter(f"{W}p"):
        runs = list(p.iter(f"{W}r"))
        ps = p.find(f"{W}pPr/{W}pStyle")
        inherited = resolve_style_ea(ps.get(f"{W}val") if ps is not None else None)
        ptext = "".join(t.text or "" for t in p.iter(f"{W}t"))
        has_break = any(b.get(f"{W}type") == "page" for b in p.iter(f"{W}br"))
        if has_break and not ptext.strip():
            empty_break += 1
        for r in runs:
            rf = r.find(f"{W}rPr/{W}rFonts")
            if rf is not None:
                for k in ("ascii", "hAnsi", "eastAsia", "cs"):
                    v = rf.get(f"{W}{k}")
                    if v:
                        fonts.add(v)
                        if k == "eastAsia":
                            ea_fonts.add(v)
            txt = "".join(t.text or "" for t in r.iter(f"{W}t"))
            if CJK.search(txt):
                ea = (rf.get(f"{W}eastAsia") if rf is not None else None) or inherited
                if not ea:
                    bad_runs.append(txt.strip()[:28])
                else:
                    ea_fonts.add(ea)

    # 图片尺寸
    images = []
    for ext in doc.iter(f"{WP}extent"):
        images.append((int(ext.get("cx", 0)), int(ext.get("cy", 0))))

    # 表格：行数、是否设了表头重复
    tables = []
    for t in doc.iter(f"{W}tbl"):
        rows = list(t.iter(f"{W}tr"))
        has_hdr = any(r.find(f"{W}trPr/{W}tblHeader") is not None for r in rows)
        tables.append({"rows": len(rows), "header_repeat": has_hdr})

    return {
        "bad_runs": bad_runs,
        "fonts": sorted(fonts | set(style_ea.values()) | ({default_ea} if default_ea else set())),
        "eastasia_fonts": sorted(ea_fonts),
        "style_eastasia": sorted(set(style_ea.values())),
        "empty_page_breaks": empty_break, "images": images, "media": len(media),
        "tables": tables, "text_width_emu": text_w_emu,
    }


def _font_installed(name: str) -> bool:
    """本机有没有这个字体。探测不了就返回 True，宁可漏报不误报。"""
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True,
                             timeout=20, text=True)
        if out.returncode == 0 and out.stdout:
            return name.lower() in out.stdout.lower()
    except Exception:
        pass
    if sys.platform == "darwin":
        try:
            out = subprocess.run(["system_profiler", "-json", "SPFontsDataType"],
                                 capture_output=True, timeout=60, text=True)
            if out.returncode == 0 and out.stdout:
                return name.lower() in out.stdout.lower()
        except Exception:
            pass
    return True


def render_pages(path: Path, shots: str | None) -> tuple[list[float] | None, str]:
    """用 soffice 转 PDF 再逐页转 PNG，算每页的内容占比。缺工具就返回 None。"""
    if not shutil.which("soffice"):
        return None, "未找到 soffice（LibreOffice），跳过空白页检查"
    if not shutil.which("pdftoppm"):
        return None, "未找到 pdftoppm（poppler），跳过空白页检查"
    try:
        from PIL import Image
    except ImportError:
        return None, "未安装 Pillow，跳过空白页检查"

    tmp = tempfile.mkdtemp(prefix="verify_docx_")
    try:
        r = subprocess.run(
            ["soffice", "--headless", f"-env:UserInstallation=file://{tmp}/profile",
             "--convert-to", "pdf", "--outdir", tmp, str(path)],
            capture_output=True, timeout=300)
        pdfs = glob.glob(os.path.join(tmp, "*.pdf"))
        if not pdfs:
            return None, f"soffice 转换失败，跳过空白页检查（{r.stderr.decode()[:80]}）"
        subprocess.run(["pdftoppm", "-png", "-r", "80", pdfs[0],
                        os.path.join(tmp, "pg")], capture_output=True, timeout=300)
        pngs = sorted(glob.glob(os.path.join(tmp, "pg-*.png")))
        ratios = []
        for f in pngs:
            im = Image.open(f).convert("L")
            w, h = im.size
            px = im.load()
            dark = sum(1 for y in range(0, h, 3) for x in range(0, w, 3) if px[x, y] < 230)
            ratios.append(dark / (len(range(0, h, 3)) * len(range(0, w, 3))))
        if shots:
            outdir = os.path.dirname(shots) or "."
            os.makedirs(outdir, exist_ok=True)
            for i, f in enumerate(pngs, 1):
                shutil.copy(f, f"{shots}_{i:02d}.png")
        return ratios, f"渲染 {len(pngs)} 页"
    except subprocess.TimeoutExpired:
        return None, "soffice 超时，跳过空白页检查"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)



# ── 自检：造三份「已知有病」的文档，断言全被拦下 ──────────────────
# 闸门自己也会烂。改过这个文件之后跑一次 --self-test，
# 确认它还抓得住它当初就是为了抓的那三类错。纯标准库，不装 python-docx。

_MIN_DOC = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>{body}'
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418"/>'
    '</w:sectPr></w:body></w:document>')

_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/'
         '2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')

_CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
       'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
       '<Override PartName="/word/document.xml" ContentType="application/vnd.'
       'openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')


def _write_docx(path: str, body: str) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("word/document.xml", _MIN_DOC.format(body=body))


def _run(text: str, ea: str | None = None, latin: str | None = None) -> str:
    rf = ""
    if ea or latin:
        parts = []
        if latin:
            parts.append(f'w:ascii="{latin}" w:hAnsi="{latin}"')
        if ea:
            parts.append(f'w:eastAsia="{ea}"')
        rf = f'<w:rPr><w:rFonts {" ".join(parts)}/></w:rPr>'
    return f'<w:p><w:r>{rf}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def self_test() -> int:
    cases = [
        ("中文 run 缺 eastAsia",
         _run("这段中文没有设置东亚字体，Word 会兜底成宋体。", latin="Calibri"),
         "没有设置 w:eastAsia"),
        ("东亚字体绑定单一平台",
         _run("苹方只有 macOS 上才有。", ea="PingFang SC"),
         "换台机器打开"),
        ("空段落里放分页符",
         _run("正文。", ea="Microsoft YaHei")
         + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
         "空段落 + 分页符"),
    ]
    tmp = tempfile.mkdtemp(prefix="verify_docx_selftest_")
    ok = True
    try:
        for i, (name, body, expect) in enumerate(cases):
            f = os.path.join(tmp, f"case{i}.docx")
            _write_docx(f, body)
            d = parse(Path(f))
            blob = json.dumps(d, ensure_ascii=False)
            hit = (expect == "没有设置 w:eastAsia" and d["bad_runs"]) or \
                  (expect == "换台机器打开" and
                   any(x in PLATFORM_ONLY for x in d["eastasia_fonts"])) or \
                  (expect == "空段落 + 分页符" and d["empty_page_breaks"])
            print(f"  {'✓' if hit else '✗'} {name}")
            ok = ok and bool(hit)
        # 反向用例：正确写法不该被误报（含「字体声明在段落样式里」这种合法写法）
        f = os.path.join(tmp, "good.docx")
        _write_docx(f, _run("这段中文显式带了东亚字体。", ea="Microsoft YaHei",
                            latin="Microsoft YaHei"))
        d = parse(Path(f))
        clean = not d["bad_runs"] and not d["empty_page_breaks"] and \
            not any(x in PLATFORM_ONLY for x in d["eastasia_fonts"])
        print(f"  {'✓' if clean else '✗'} 正确写法不误报")
        ok = ok and clean
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("\n自检" + ("通过" if ok else "未通过——闸门坏了，先修它再用它验别的文档"))
    return 0 if ok else 1


def main() -> int:
    ap = argparse.ArgumentParser(description="Word 文档交付前自检")
    ap.add_argument("docx", nargs="?")
    ap.add_argument("--self-test", action="store_true",
                    help="造三份已知有病的文档，确认闸门还抓得住")
    ap.add_argument("--shots", metavar="前缀", help="额外把每页导出成 PNG")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--min-ink", type=float, default=0.02,
                    help="内容占比低于此值判为近乎空白页，默认 0.02")
    a = ap.parse_args()

    if a.self_test:
        return self_test()
    if not a.docx:
        ap.error("需要给一个 .docx 路径，或用 --self-test")

    path = Path(a.docx)
    if not path.exists():
        print(f"找不到文件：{path}", file=sys.stderr)
        return 2

    # ⓪ 包结构：坏包一票否决，后面的检查都建立在能解析之上
    health = xml_health(path)
    if health:
        print("❌ FAIL：包结构已损坏，其余检查跳过")
        for h in health:
            print("   -", h)
        return 1

    d = parse(path)
    fails: list[str] = []
    warns: list[str] = []
    skipped: list[str] = []

    # ① 中文 run 缺 eastAsia —— 整篇字体走样的头号成因
    if d["bad_runs"]:
        n = len(d["bad_runs"])
        uniq = list(dict.fromkeys(d["bad_runs"]))
        sample = "、".join(f"「{s}」" for s in uniq[:3])
        fails.append(
            f"{n} 个含中文的 run 没有设置 w:eastAsia，Word 会用主题字体兜底"
            f"（通常回退成宋体）。样例：{sample}")

    # ② 跨平台风险字体
    #    中文字体绑定单一平台 → FAIL：整篇中文走样，且换机器之前完全看不出来。
    #    西文字体绑定单一平台 → WARN：Word 会回退到度量相近的字体，后果轻。
    ea_declared = {v for v in d["eastasia_fonts"]}
    for fnt in sorted(ea_declared):
        if fnt in PLATFORM_ONLY:
            fails.append(
                f"东亚字体「{fnt}」{PLATFORM_ONLY[fnt]}——换台机器打开，"
                f"整篇中文会回退成主题字体。改用 Office 双平台自带的："
                f"Microsoft YaHei / DengXian / SimSun / SimHei")
    for fnt in sorted(set(d["fonts"]) - ea_declared):
        if fnt in PLATFORM_ONLY:
            warns.append(f"西文字体「{fnt}」{PLATFORM_ONLY[fnt]}——换平台会回退到度量相近的字体")
    if ea_declared and not (ea_declared & SAFE_EASTASIA) and not (ea_declared & set(PLATFORM_ONLY)):
        warns.append(
            f"东亚字体「{'、'.join(sorted(ea_declared))}」不在 Office 双平台自带清单里，"
            "换机器可能变样。安全清单：Microsoft YaHei / DengXian / SimSun / SimHei / KaiTi / FangSong")

    # ③ 空段落里的分页符 —— 整页空白的成因
    if d["empty_page_breaks"]:
        warns.append(
            f"{d['empty_page_breaks']} 处「空段落 + 分页符」。空段落本身占一行，"
            "上一页正好排满时会被挤到下一页、分页符再把正文推到再下一页，"
            "于是留下一整页全白。改用段落属性 page_break_before。")

    # ④ 图片宽过版心
    if d["text_width_emu"]:
        over = [(cx, cy) for cx, cy in d["images"] if cx > d["text_width_emu"]]
        if over:
            fails.append(
                f"{len(over)} 张图片宽度超出版心"
                f"（版心 {d['text_width_emu']/EMU_PER_CM:.1f}cm，"
                f"最宽的图 {max(c for c,_ in over)/EMU_PER_CM:.1f}cm），右侧会被裁掉")

    # ⑤ 长表格没设表头重复
    for i, t in enumerate(d["tables"], 1):
        if t["rows"] >= 12 and not t["header_repeat"]:
            warns.append(f"表 {i} 有 {t['rows']} 行且未设置表头重复"
                         "（w:tblHeader），跨页后读者看不到列名")

    # ⑥ 一张图都没有
    if not d["images"]:
        warns.append("文档里没有任何图片。图表被降级成数据表时会这样——"
                     "确认这是有意的取舍，而不是图没嵌进去")

    # ⑦ 空白页（需要渲染）
    ratios, note = render_pages(path, a.shots)
    pages = None
    if ratios is None:
        skipped.append(note)
    else:
        pages = len(ratios)
        missing = [f for f in d["fonts"] if not _font_installed(f)]
        if missing:
            skipped.append(
                f"渲染时本机缺字体：{'、'.join(missing)}。LibreOffice 用替代字体排的版，"
                "所以页数与空白页位置只是近似——最终请在 Word 里再看一眼")
        blanks = [i + 1 for i, r in enumerate(ratios[:-1]) if r < a.min_ink]
        if blanks:
            where = f"第 {'、'.join(map(str, blanks))} 页近乎空白（末页除外，末页短属正常）"
            if missing:
                # 字体都缺了，分页本来就是近似的——不能拿这个证据下硬结论。
                warns.append(where + "。但本机缺字体、排版是替代字体算出来的，"
                             "这个页码不一定对得上 Word 里的实际情况，请在 Word 里复核")
            else:
                fails.append(where)

    if a.json:
        print(json.dumps({"fails": fails, "warns": warns, "skipped": skipped,
                          "pages": pages, "images": len(d["images"]),
                          "tables": len(d["tables"]), "fonts": d["fonts"]},
                         ensure_ascii=False, indent=2))
        return 1 if fails else 0

    line = "─" * 68
    print(line)
    head = f"Word 自检：{path.name}"
    if pages:
        head += f"   {pages} 页"
    head += f" / {len(d['images'])} 图 / {len(d['tables'])} 表"
    print(head)
    print(line)
    if fails:
        print(f"\n❌ FAIL {len(fails)} 项（必须修）")
        for x in fails:
            print(f"   - {x}")
    if warns:
        print(f"\n⚠ WARN {len(warns)} 项（自己判断）")
        for x in warns:
            print(f"   - {x}")
    if skipped:
        print(f"\n○ 跳过 {len(skipped)} 项（工具缺失，**不等于通过**）")
        for x in skipped:
            print(f"   - {x}")
    if not fails and not warns:
        if skipped:
            print(f"\n✅ 已验的项目全部通过——但上面 {len(skipped)} 项没验成，"
                  "别把它当作全部通过")
        else:
            print("\n✅ 全部通过")
    print(f"\n用到的字体：{'、'.join(d['fonts']) or '（未显式指定，全走主题字体）'}")
    print("\n机器只判「有没有画错」。这几件它判不了，自己翻一遍导出的页面图：")
    print("  · 每张图能不能独立回答它标题里那句话")
    print("  · 表格有没有该合并却拆开的列")
    print("  · 首页能不能在十秒内看懂核心结论")
    print(line)
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
