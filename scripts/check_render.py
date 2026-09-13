# -*- coding: utf-8 -*-
"""渲染校验器：把 PPTX 渲染成 PNG，输出①字符画预览 ②彩色元素位置报告。

设计动机：AI 无法直接"看"图片，几何自检又只能验证数值关系。
字符画预览把布局转成可读文本，一眼即可判断箭头是否飘、元素是否错位。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptxfig import export_png
from PIL import Image

# 颜色 → 字符（按亮度/色相分类）。顺序敏感：先判白底，再判其他。
def _is_white(r, g, b):
    return r > 238 and g > 238 and b > 238


PALETTE = [
    (".", _is_white),                                                    # 白底
    ("R", lambda r, g, b: r > 120 and g < 115 and b < 115),              # 红：箭头/里程碑
    ("D", lambda r, g, b: b > 80 and r < 100 and g < 140),               # 深蓝：中心框
    ("B", lambda r, g, b: b > 130 and r < 150 and g > 95 and g < 200),   # 中蓝：连线
    ("o", lambda r, g, b: b > 195 and r > 150 and g > 170
     and (r + g + b) < 735),                                             # 浅蓝：图框填充
    ("*", lambda r, g, b: r < 150 and g < 150 and b < 150),              # 深灰：文字
    ("-", lambda r, g, b: r > 195 and g > 195 and b > 195),              # 浅灰：网格线
]


def classify(r, g, b):
    for ch, pred in PALETTE:
        try:
            if pred(r, g, b):
                return ch
        except Exception:
            pass
    return "+"


def ascii_preview(png, cols=78, rows=30, w_in=13.333):
    """把渲染图转成字符画。
    关键点：不要在缩放图上平均采样（细线/文字会被抹掉），
    而是对每个字符格取「最偏离白色」的像素，保证细元素可见。
    """
    img = Image.open(png).convert("RGB")
    px = img.load()
    W, H = img.size
    bw, bh = W / float(cols), H / float(rows)
    lines = []
    for ry in range(rows):
        row = []
        for rx in range(cols):
            x0, x1 = int(rx * bw), max(int(rx * bw) + 1, int((rx + 1) * bw))
            y0, y1 = int(ry * bh), max(int(ry * bh) + 1, int((ry + 1) * bh))
            best, best_dev = None, -1
            for y in range(y0, min(y1, H)):
                for x in range(x0, min(x1, W)):
                    r, g, b = px[x, y][:3]
                    dev = 255 * 3 - (r + g + b)          # 偏离白色的程度
                    if dev > best_dev:
                        best_dev, best = dev, (r, g, b)
            row.append(classify(*best) if best else ".")
        lines.append("".join(row))
    return lines, img.size


def color_blocks(png, pred, min_px=8, w_in=13.333):
    """找出满足颜色条件的像素团，返回每个团的 (中心x_in, 中心y_in, 像素数)。
    简易连通域：按网格聚类（足够用于定位箭头）。"""
    img = Image.open(png).convert("RGB")
    px = img.load()
    W, H = img.size
    ppi = W / w_in
    pts = []
    for y in range(0, H, 2):
        for x in range(0, W, 2):
            r, g, b = px[x, y][:3]
            if pred(r, g, b):
                pts.append((x, y))
    if not pts:
        return []
    # 网格聚类（8px）
    cell = 8
    grid = {}
    for x, y in pts:
        grid.setdefault((x // cell, y // cell), []).append((x, y))
    # 合并相邻格
    keys = set(grid)
    clusters = []
    for k in sorted(keys):
        merged = None
        for c in clusters:
            if any((abs(k[0] - g[0]) <= 1 and abs(k[1] - g[1]) <= 1) for g in c):
                c.append(k)
                merged = c
                break
        if merged is None:
            clusters.append([k])
    out = []
    for c in clusters:
        allp = [p for g in c for p in grid[g]]
        if len(allp) < min_px:
            continue
        cx = sum(p[0] for p in allp) / len(allp) / ppi
        cy = sum(p[1] for p in allp) / len(allp) / ppi
        out.append((cx, cy, len(allp)))
    return sorted(out, key=lambda t: (round(t[1], 1), t[0]))


IS_RED = lambda r, g, b: r > 130 and g < 110 and b < 110


def report(pptx_path, tmp_dir=None, max_pages=3, show_ascii=True):
    tmp_dir = tmp_dir or os.path.join(os.path.dirname(pptx_path), "_render_tmp")
    pngs = export_png(pptx_path, tmp_dir)
    for i, png in enumerate(pngs[:max_pages], 1):
        print(f"\n{'=' * 78}\n第 {i} 页  {os.path.basename(png)}")
        if show_ascii:
            lines, size = ascii_preview(png)
            print(f"[{size[0]}x{size[1]}]  图例: R=红(箭头) D=深蓝 B=中蓝 o=浅蓝框 *=文字 -=浅灰")
            print("+" + "-" * len(lines[0]) + "+")
            for ln in lines:
                print("|" + ln + "|")
            print("+" + "-" * len(lines[0]) + "+")
        blocks = color_blocks(png, IS_RED)
        if blocks:
            print(f"红色元素（箭头/里程碑）{len(blocks)} 处：")
            for cx, cy, n in blocks[:14]:
                print(f"    x={cx:5.2f}in  y={cy:5.2f}in  像素{n}")
        else:
            print("红色元素：无")
    return pngs


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python check_render.py <pptx路径> [临时目录]")
        sys.exit(1)
    report(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
