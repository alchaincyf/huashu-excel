#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_visual.py — 交付前的渲染闸门。

数字有 verify_numbers.py 把关，图表一直没有。手写 SVG 必然会犯的那几类错
（线画出框、标签叠成一团、文字压住柱子、图上数字和正文对不上），
肉眼扫一遍看不出来，渲染出来一看全在。这个脚本把它们变成退出码。

它只判机器能判死的事实，不判美不美——那是你的活。

用法：
    python3 verify_visual.py 报告.html
    python3 verify_visual.py 报告.html --shots 自检      # 顺带存分屏截图
    python3 verify_visual.py 报告.html --json

退出码：0 = 通过；1 = 有 FAIL；2 = 环境缺 playwright（不算失败，会提示改用人工）

依赖 playwright（pip install playwright && playwright install chromium）。
装不上就手动开一眼浏览器 —— 但别跳过这一步。
"""
import sys, os, json, argparse, re

CHECKS = """() => {
  const out = {overflow:null, oob:[], textOverlap:[], textOnMark:[], tiny:[],
               svgCount:0, dualAxis:[], numbers:{svg:[], body:[]}};

  // 1) 页面横向溢出
  const de = document.documentElement;
  out.overflow = {scrollW: de.scrollWidth, clientW: de.clientWidth,
                  overflows: de.scrollWidth > de.clientWidth + 1};

  const rectsOverlap = (a,b) => !(a.x+a.width <= b.x || b.x+b.width <= a.x ||
                                  a.y+a.height <= b.y || b.y+b.height <= a.y);
  const area = r => Math.max(0,r.width)*Math.max(0,r.height);
  const inter = (a,b) => {
    const x = Math.max(a.x,b.x), y = Math.max(a.y,b.y);
    const w = Math.min(a.x+a.width, b.x+b.width) - x;
    const h = Math.min(a.y+a.height, b.y+b.height) - y;
    return (w>0&&h>0) ? w*h : 0;
  };

  const svgs = [...document.querySelectorAll('svg')];
  out.svgCount = svgs.length;

  svgs.forEach((svg, si) => {
    const vb = svg.viewBox.baseVal;
    const W = vb && vb.width ? vb.width : svg.getBBox().width;
    const H = vb && vb.height ? vb.height : svg.getBBox().height;
    const label = svg.getAttribute('aria-label') || svg.getAttribute('data-name') || ('svg#'+si);
    const pad = 2;   // 允许描边等微小外扩

    const texts = [];
    [...svg.querySelectorAll('text, tspan, rect, circle, polyline, polygon, line, path')].forEach(el => {
      let bb; try { bb = el.getBBox(); } catch(e) { return; }
      if (!bb || (!bb.width && !bb.height)) return;

      // 2) 越出 viewBox
      if (bb.x < -pad || bb.y < -pad || bb.x+bb.width > W+pad || bb.y+bb.height > H+pad) {
        out.oob.push({svg: label, tag: el.tagName,
          text: (el.textContent||'').trim().slice(0,28),
          box: [+bb.x.toFixed(1), +bb.y.toFixed(1), +bb.width.toFixed(1), +bb.height.toFixed(1)],
          viewBox: [W, H]});
      }
      if (el.tagName === 'text') {
        const t = (el.textContent||'').trim();
        if (t) texts.push({el, bb, t});
      }
    });

    // 3) 文字之间互相重叠
    for (let i=0;i<texts.length;i++) for (let j=i+1;j<texts.length;j++) {
      const A=texts[i], B=texts[j];
      if (!rectsOverlap(A.bb,B.bb)) continue;
      const ov = inter(A.bb,B.bb);
      const frac = ov / Math.min(area(A.bb), area(B.bb));
      if (frac > 0.18) out.textOverlap.push({svg: label, a: A.t.slice(0,24), b: B.t.slice(0,24),
                                            overlapPct: +(frac*100).toFixed(0)});
    }

    // 4) 文字压在图形标记上（柱/点/线），只报遮挡明显的
    const marks = [...svg.querySelectorAll('rect, circle, polygon')].map(el => {
      let bb; try { bb = el.getBBox(); } catch(e) { return null; }
      if (!bb || area({x:0,y:0,width:bb.width,height:bb.height}) < 40) return null;
      const op = parseFloat(el.getAttribute('opacity') || getComputedStyle(el).opacity || '1');
      const fill = (el.getAttribute('fill')||'').toLowerCase();
      if (op < 0.25 || fill === 'none') return null;   // 半透明底纹不算遮挡
      return {el, bb};
    }).filter(Boolean);
    texts.forEach(T => {
      // 累计所有图形对这行文字的遮挡：一行长标签横穿多根柱子时，
      // 单根柱子占比都不高，但合起来已经压得看不清了。
      // 只算「画在文字之上」的图形——SVG 按文档顺序绘制，排在文字之后的才会盖住它。
      // 白字写在色块上是正常设计，那种情况色块在前、文字在后，不算遮挡。
      let covered = 0, tags = new Set();
      marks.forEach(M => {
        if (M.el === T.el) return;
        const pos = T.el.compareDocumentPosition(M.el);
        const markIsAbove = !!(pos & Node.DOCUMENT_POSITION_FOLLOWING);
        if (!markIsAbove) return;
        const ov = inter(T.bb, M.bb);
        if (ov > 0) { covered += ov; tags.add(M.el.tagName); }
      });
      const frac = covered / Math.max(1, area(T.bb));
      if (frac > 0.35)
        out.textOnMark.push({svg: label, text: T.t.slice(0,28),
                             mark: [...tags].join('/'), coverPct: +(frac*100).toFixed(0)});
    });

    // 5) 疑似双轴：左右两侧各有一组刻度文字
    const leftTicks = texts.filter(t => t.bb.x < W*0.12 && /^[\\d,.\\-–%万k ]+$/.test(t.t));
    const rightTicks = texts.filter(t => t.bb.x+t.bb.width > W*0.88 && /^[\\d,.\\-–%万k ]+$/.test(t.t));
    if (leftTicks.length >= 3 && rightTicks.length >= 3)
      out.dualAxis.push({svg: label, left: leftTicks.length, right: rightTicks.length});

    // 6) 渲染尺寸过小（图被压扁基本等于没有）
    const r = svg.getBoundingClientRect();
    if (r.width < 200 || r.height < 60)
      out.tiny.push({svg: label, w: +r.width.toFixed(0), h: +r.height.toFixed(0)});

    // 收集 SVG 内的数字，供与正文交叉核对
    texts.forEach(t => {
      const m = t.t.replace(/,/g,'').match(/-?\\d+(?:\\.\\d+)?/g);
      if (m) m.forEach(x => out.numbers.svg.push(parseFloat(x)));
    });
  });

  // 正文数字（排除 svg 内的）
  const clone = document.body.cloneNode(true);
  [...clone.querySelectorAll('svg')].forEach(s => s.remove());
  const bodyText = clone.textContent || '';
  const bm = bodyText.replace(/,/g,'').match(/-?\\d+(?:\\.\\d+)?/g);
  if (bm) bm.forEach(x => out.numbers.body.push(parseFloat(x)));

  return out;
}"""


# 量之前先把内容全亮出来。两类结构性误报的根源都是「量了看不见的东西」：
# ① 幻灯片用 display:none 藏非当前页，隐藏元素的包围盒必然 0×0，
#    于是任何 deck 都 100% FAIL「渲染尺寸过小」（04 号压测把病因误诊成 CSS）；
# ② <details> 折叠的附录同理。
# 展开不改变任何几何——只是让本来要人翻到才看得见的内容变成可测量的。
EXPAND = """() => {
  document.querySelectorAll('details:not([open])').forEach(d => d.open = true);
  let n = 0;
  document.querySelectorAll('body *').forEach(e => {
    if (getComputedStyle(e).display === 'none' && e.querySelector('svg')) {
      e.style.setProperty('display', 'block', 'important'); n++;
    }
  });
  return n;
}"""


def near_miss(svg_nums, body_nums, rel=0.008, min_abs=100):
    """图上数字与正文数字「接近但不相等」——多半是两处用了不同的取整口径。
    阈值刻意收得很紧（0.8%）：宁可漏报，也不要用一堆巧合淹掉真问题。"""
    body = sorted(set(b for b in body_nums if abs(b) >= min_abs))
    hits = []
    seen = set()
    for s in sorted(set(svg_nums)):
        if abs(s) < min_abs:
            continue
        if s in body:
            continue
        best = None
        for b in body:
            if b == s:
                continue
            d = abs(b - s) / max(abs(s), 1e-9)
            if d <= rel and (best is None or d < best[1]):
                best = (b, d)
        if best and (s, best[0]) not in seen:
            seen.add((s, best[0]))
            hits.append({'svg': s, 'body': best[0], 'diffPct': round(best[1] * 100, 2)})
    return hits


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('html')
    ap.add_argument('--shots', metavar='前缀', help='额外保存分屏截图')
    ap.add_argument('--width', type=int, default=None,
                    help='视口宽。默认：报告 1100，deck 自动切 1920')
    ap.add_argument('--deck', action='store_true',
                    help='按幻灯片验：多页隐藏结构会被自动识别，此开关用于强制')
    ap.add_argument('--json', action='store_true')
    ap.add_argument('--allow-dual-axis', action='store_true',
                    help='确实需要双轴且已在图上标注刻度为人为选定时使用')
    a = ap.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print('⚠ 未安装 playwright，无法自动验图。', file=sys.stderr)
        print('  装：pip install playwright && playwright install chromium', file=sys.stderr)
        print('  不装就手动在浏览器里逐屏看一遍——这一步不能跳过。', file=sys.stderr)
        sys.exit(2)

    path = os.path.abspath(a.html)
    if not os.path.exists(path):
        print(f'找不到文件：{path}', file=sys.stderr); sys.exit(1)

    errs = []
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={'width': a.width or 1100, 'height': 1000},
                        device_scale_factor=2)
        pg.on('console', lambda m: errs.append(f'[{m.type}] {m.text}') if m.type == 'error' else None)
        pg.on('pageerror', lambda e: errs.append(f'[pageerror] {e}'))
        pg.goto('file://' + path)
        pg.wait_for_timeout(900)
        expanded = pg.evaluate(EXPAND)
        is_deck = a.deck or expanded >= 2
        if is_deck and a.width is None:
            # deck 按投影尺寸设计，拿 1100 宽的报告视口量它，
            # 「页面横向溢出」这类读数全是视口错配的产物（06 号压测实锤）
            pg.set_viewport_size({'width': 1920, 'height': 1080})
            pg.wait_for_timeout(300)
        R = pg.evaluate(CHECKS)
        H = pg.evaluate('document.documentElement.scrollHeight')
        if a.shots:
            slides = pg.query_selector_all('.slide') if is_deck else []
            if slides:
                for i, s in enumerate(slides):
                    # 逐页强制可见再截：纯文字页（口径/边界）没有 svg，
                    # 不在 EXPAND 的展开范围里，不亮出来截不到
                    s.evaluate("el => el.style.setProperty('display','flex','important')")
                    s.scroll_into_view_if_needed()
                    pg.wait_for_timeout(120)
                    s.screenshot(path=f'{a.shots}_{i}.png')
            else:
                y = 0; i = 0
                while y < H and i < 24:
                    pg.evaluate(f'window.scrollTo(0,{y})'); pg.wait_for_timeout(180)
                    pg.screenshot(path=f'{a.shots}_{i}.png'); y += 950; i += 1
                pg.screenshot(path=f'{a.shots}_整页.png', full_page=True)
        b.close()

    R['consoleErrors'] = errs
    R['nearMiss'] = near_miss(R['numbers']['svg'], R['numbers']['body'])
    R.pop('numbers', None)

    R['expanded'] = expanded
    R['isDeck'] = is_deck
    if a.json:
        print(json.dumps(R, ensure_ascii=False, indent=1));
        sys.exit(1 if (R['oob'] or R['textOverlap'] or R['textOnMark'] or R['tiny']
                       or (R['overflow']['overflows'] and not is_deck) or errs) else 0)

    fails, warns = [], []
    if R['overflow']['overflows']:
        msg = f"页面横向溢出：scrollWidth {R['overflow']['scrollW']} > 视口 {R['overflow']['clientW']}"
        if is_deck:
            warns.append(msg + "（deck 常为固定页宽设计，先确认设计宽度再下判断）")
        else:
            fails.append(msg)
    for o in R['oob']:
        fails.append(f"[{o['svg']}] <{o['tag']}> 画出 viewBox：box={o['box']} viewBox={o['viewBox']}"
                     + (f" 内容「{o['text']}」" if o['text'] else ''))
    for o in R['textOverlap']:
        fails.append(f"[{o['svg']}] 标签重叠 {o['overlapPct']}%：「{o['a']}」×「{o['b']}」")
    for o in R['textOnMark']:
        fails.append(f"[{o['svg']}] 文字被 <{o['mark']}> 遮挡 {o['coverPct']}%：「{o['text']}」")
    for o in R['tiny']:
        fails.append(f"[{o['svg']}] 渲染尺寸过小 {o['w']}×{o['h']}px")
    for e in errs:
        fails.append(f"控制台错误：{e}")
    for o in R['dualAxis']:
        (warns if a.allow_dual_axis else fails).append(
            f"[{o['svg']}] 疑似双 Y 轴（左 {o['left']} 个刻度 / 右 {o['right']} 个）。"
            "双轴的视觉对比取决于你选的两个缩放比例，换个比例就换个结论——"
            "优先改成指数化（同起点=100）或上下两张共享 X 轴的小图。"
            "确有必要请加 --allow-dual-axis 并在图上注明刻度为人为选定。")
    for o in R['nearMiss']:
        warns.append(f"图上 {o['svg']} 与正文 {o['body']} 相差 {o['diffPct']}%——"
                     "同一个量两处取整口径可能不一致，核一下")

    print('─' * 68)
    print(f'渲染自检：{os.path.basename(path)}   共 {R["svgCount"]} 张 SVG，页高 {H}px')
    if expanded:
        print(f'（已先展开 {expanded} 个隐藏容器再测量'
              + ('，按 deck 模式验，视口 1920×1080' if is_deck else '') + '）')
    print('─' * 68)
    if fails:
        print(f'\n❌ FAIL {len(fails)} 项（必须修）')
        for f in fails: print('   -', f)
    if warns:
        print(f'\n⚠ WARN {len(warns)} 项（自己判断）')
        for w in warns: print('   -', w)
    if not fails and not warns:
        print('\n✅ 机器能查的都通过了。')
    print('\n机器只判「画错了没有」。这几件它判不了，你自己看一眼截图：')
    print('  · 每张图能不能独立回答标题里那句话')
    print('  · 有没有该画成图却写成了段落的地方')
    print('  · 首屏能不能在 10 秒内看懂核心结论')
    print('─' * 68)
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
