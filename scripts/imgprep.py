#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把凭证原件（PDF / 扫描图 / 手机照片）处理成 8cm 宽的佐证图片。

处理链：渲染 → 裁白边 → 画红框 → 缩放到 8cm(605px 宽) → 输出 <tag>_01.png …

用法：
  # PDF 全页
  python3 imgprep.py 立项通知.pdf --out ./img --tag m1

  # 单张图
  python3 imgprep.py 奖状.jpg --out ./img --tag m1

  # 先看表格横竖线位置，再决定红框坐标
  python3 imgprep.py 名单.pdf --out ./img --tag m1 --detect

  # 框住某个表格行（坐标以"裁边后"的图片为准，用 --detect 得到）
  python3 imgprep.py 名单.pdf --out ./img --tag m1 --box 1127:1159

  # 任意矩形
  python3 imgprep.py 通知.jpg --out ./img --tag m2 --rect 75,1057,1164,1221

说明：
  - 红框以外不要做任何涂改；原件的红头、公章、文号必须完整保留。
  - --box/--rect 可重复，用于同时框多处。
  - 输出宽度固定 605px，在 Word 中按 8cm 显示约 192dpi，打印清晰。
"""

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys

from PIL import Image, ImageChops, ImageDraw

OUT_W = 605          # 8cm 宽对应像素（≈192dpi）
PAD = 10             # 裁边后留白
BOX_COLOR = (255, 0, 0)
BOX_WIDTH = 4


def render_pdf(pdf, outdir, dpi):
    """用 ghostscript 渲染 PDF 全部页面。"""
    gs = shutil.which("gs")
    if not gs:
        raise SystemExit("未找到 ghostscript（gs），无法渲染 PDF")
    os.makedirs(outdir, exist_ok=True)
    pattern = os.path.join(outdir, "page_%03d.png")
    subprocess.run([gs, "-q", "-dNOPAUSE", "-dBATCH", "-sDEVICE=png16m",
                    "-r%d" % dpi, "-sOutputFile=" + pattern, pdf],
                   check=True, capture_output=True)
    return sorted(glob.glob(os.path.join(outdir, "page_*.png")))


def trim(im, pad=PAD):
    """裁掉四周空白，返回 (图, (left,top))。"""
    rgb = im.convert("RGB")
    gray = rgb.convert("L")
    bg = Image.new("L", gray.size, 255)
    diff = ImageChops.difference(gray, bg).point(lambda v: 255 if v > 12 else 0)
    box = diff.getbbox()
    if not box:
        return rgb, (0, 0)
    l, t, r, b = box
    l2, t2 = max(0, l - pad), max(0, t - pad)
    return rgb.crop((l2, t2, min(rgb.width, r + pad), min(rgb.height, b + pad))), (l2, t2)


def detect_lines(im, thr=205, ratio=0.55):
    """检测表格横线/竖线，打印坐标（裁边后坐标，可直接用于 --box / --rect）。"""
    g = im.convert("L")
    w, h = g.size
    px = g.load()

    def runs(values):
        groups = []
        for v in values:
            if groups and v - groups[-1][-1] <= 2:
                groups[-1].append(v)
            else:
                groups.append([v])
        return [sum(gp) // len(gp) for gp in groups]

    ys = []
    for y in range(0, h, 1):
        dark = sum(1 for x in range(20, w - 20, 2) if px[x, y] < thr)
        if dark > ((w - 40) / 2) * ratio:
            ys.append(y)
    xs = []
    for x in range(0, w, 1):
        dark = sum(1 for y in range(20, h - 20, 2) if px[x, y] < thr)
        if dark > ((h - 40) / 2) * ratio:
            xs.append(x)
    hy, vx = runs(ys), runs(xs)
    print("尺寸 %dx%d" % (w, h))
    print("横线 y = %s" % hy)
    print("竖线 x = %s" % vx)
    if len(hy) >= 2:
        print("相邻行高 = %s" % [hy[i + 1] - hy[i] for i in range(len(hy) - 1)])
    print("\n框某一行：--box %s:%s（取相邻两条横线，可各外扩 2px）" %
          (hy[-2] if len(hy) >= 2 else "y0", hy[-1] if hy else "y1"))


def pad_to_frame(im, ratio=856 / 605.0):
    """把图放进固定宽高比的白色画框（内容等比缩放、顶部对齐、水平居中）。

    多页公文务必用这个：逐页等比会导致每页高度不一致，
    两行图加标题刚好卡在版心边缘时，会时好时坏（一页 4 张 / 2 张跳变）。
    """
    w, h = im.size
    tw = max(w, int(round(h / ratio)))
    th = int(round(tw * ratio))
    if th < h:
        tw = int(round(h / ratio))
        th = int(round(tw * ratio))
    canvas = Image.new("RGB", (tw, th), "white")
    canvas.paste(im, ((tw - w) // 2, 0))
    return canvas


def draw_boxes(im, boxes, rects):
    d = ImageDraw.Draw(im)
    w, h = im.size
    for y0, y1 in boxes:
        d.rectangle([2, y0, w - 3, y1], outline=BOX_COLOR, width=BOX_WIDTH)
    for x0, y0, x1, y1 in rects:
        d.rectangle([x0, y0, x1, y1], outline=BOX_COLOR, width=BOX_WIDTH)
    return im


def parse_box(s):
    a, b = s.split(":")
    return int(a), int(b)


def parse_rect(s):
    parts = [int(v) for v in s.split(",")]
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("矩形格式应为 x0,y0,x1,y1")
    return tuple(parts)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--out", required=True)
    ap.add_argument("--tag", default="m1")
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--detect", action="store_true", help="只打印表格线坐标")
    ap.add_argument("--box", action="append", default=[], type=parse_box, help="横向通栏红框 y0:y1")
    ap.add_argument("--rect", action="append", default=[], type=parse_rect, help="矩形红框 x0,y0,x1,y1")
    ap.add_argument("--no-trim", action="store_true")
    ap.add_argument("--frame", action="store_true",
                    help="套 A4 比例白色画框（多页公文建议开启，避免各页高度不一）")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    src = args.input
    tmpdir = os.path.join(args.out, "_raw")
    if src.lower().endswith(".pdf"):
        pages = render_pdf(src, tmpdir, args.dpi)
    else:
        pages = [src]
    if not pages:
        raise SystemExit("没有可处理的页面")

    for i, page in enumerate(pages, 1):
        im = Image.open(page).convert("RGB")
        if not args.no_trim:
            im, _ = trim(im)
        if args.detect:
            print("== 第 %d 页 %s ==" % (i, os.path.basename(page)))
            detect_lines(im)
            continue
        if args.box or args.rect:
            im = draw_boxes(im, args.box, args.rect)
        if args.frame:
            im = pad_to_frame(im)
        w, h = im.size
        new_h = max(1, int(round(OUT_W * h / float(w))))
        im = im.resize((OUT_W, new_h), Image.LANCZOS)
        dst = os.path.join(args.out, "%s_%02d.png" % (args.tag, i))
        im.save(dst)
        print("%s  %dx%d  (8cm 宽，高 %.1fcm)" % (dst, OUT_W, new_h, new_h / OUT_W * 8))
    if not args.detect and len(pages) > 1 and os.path.isdir(tmpdir):
        shutil.rmtree(tmpdir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
