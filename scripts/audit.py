#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批量复核佐证材料 docx（纯 OOXML 解析，不打开编辑器、不依赖浏览器）。

用法：
  python3 audit.py 目录
  python3 audit.py 目录 --expect 任务D:6     # 指定某任务应有 6 个材料
  python3 audit.py 目录 --content             # 追加内容口径提示（不计入问题）

检查项：抬头/年度/任务标题、描述字数、清单行数与序号、材料标题字号与颜色、
图片数量与宽度/变形/同行等高、cantSplit（兼容 w:val="1" 变体）、模板占位残留、
页边距、文件完整性。--content 另查内部归属声明、描述虚报、年份与数量词口径。
"""

import argparse
import glob
import os
import re
import sys
import zipfile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def text_of(xml):
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def load(path):
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        doc = z.read("word/document.xml").decode("utf-8")
        styles = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in z.namelist() else ""
        media = [n for n in z.namelist() if n.startswith("word/media/") and not n.endswith("/")]
    return bad, doc, styles, media


def check(path, content=False):
    bad, doc, styles, media = load(path)
    paras = [p for p in re.findall(r"<w:p(?: [^>]*)?>.*?</w:p>", doc, re.S)]
    filled = [p for p in paras if text_of(p).strip()]
    info = {"file": os.path.basename(path), "problems": [], "notes": []}

    # 抬头 / 年度 / 任务标题 / 描述
    info["group"] = text_of(filled[0]) if filled else ""
    info["year_line"] = text_of(filled[1]) if len(filled) > 1 else ""
    info["task"] = text_of(filled[2]) if len(filled) > 2 else ""
    desc = text_of(filled[3]) if len(filled) > 3 else ""
    info["desc_len"] = len(re.sub(r"\s", "", desc))

    if bad:
        info["problems"].append("zip 损坏")
    for kw in ("XXXX", "写一段话描述", "竖图宽度", "横图宽度", "图片相关文字部分", "XXX专业群"):
        if kw in text_of(doc):
            info["problems"].append("占位残留：%s" % kw)
    if re.search(r"任务\s*\d-\d{1,2}-\d{1,3}\s*XXXX", text_of(doc)):
        info["problems"].append("占位残留：任务编号+XXXX")
    if not re.match(r"^\d{4}年度任务佐证材料$", info["year_line"].strip()):
        info["problems"].append("年度行异常：%s" % info["year_line"][:20])
    if not 100 <= info["desc_len"] <= 150:
        info["problems"].append("描述 %d 字（应 100-150）" % info["desc_len"])

    # 材料标题
    hs = [p for p in paras if re.match(r"^材料[一二三四五六七八九十]", text_of(p))]
    info["materials"] = len(hs)
    for p in hs:
        t = text_of(p)
        sz = re.findall(r'<w:sz w:val="(\d+)"/>', p)
        if sz and int(sz[0]) != 32:
            info["problems"].append("%s 字号 %gpt（应 16）" % (t[:8], int(sz[0]) / 2))

    # 清单表
    tbls = re.findall(r"<w:tbl>.*?</w:tbl>", doc, re.S)
    if tbls:
        rows = re.findall(r"<w:tr(?: [^>]*)?>.*?</w:tr>", tbls[0], re.S)
        info["list_rows"] = len(rows) - 1
        if len(rows) - 1 != info["materials"]:
            info["problems"].append("清单 %d 行 ≠ 材料 %d 个" % (len(rows) - 1, info["materials"]))
        for i, tr in enumerate(rows[1:], 1):
            tcs = re.findall(r"<w:tc>.*?</w:tc>", tr, re.S)
            if tcs and text_of(tcs[0]).strip() != str(i):
                info["problems"].append("清单第 %d 行序号为 %r" % (i, text_of(tcs[0]).strip()))

    # 图片
    ext = re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc)
    widths = sorted({round(int(a) / 360000, 2) for a, _ in ext})
    info["images"] = len(ext)
    info["widths"] = widths
    if media and len(media) != len(ext):
        info["problems"].append("媒体 %d 个 ≠ 引用 %d 处" % (len(media), len(ext)))
    for w in widths:
        if abs(w - 8.0) > 0.06:
            info["problems"].append("图片宽 %.2fcm（应 8cm）" % w)
    centered = sum(1 for p in paras if "<wp:inline" in p and 'w:jc w:val="center"' in p)
    if ext and centered != len(ext):
        info["problems"].append("图片居中 %d/%d" % (centered, len(ext)))

    # 变形检查：显示宽高比必须等于原图宽高比（否则文字被压扁/拉长）
    try:
        import io
        from PIL import Image
        with zipfile.ZipFile(path) as z2:
            relmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="media/([^"]+)"',
                                     z2.read("word/_rels/document.xml.rels").decode("utf-8")))
            emb = re.findall(r'r:embed="([^"]+)"', doc)
            bad = []
            for i, ((cx, cy), rid) in enumerate(zip(ext, emb), 1):
                if rid not in relmap:
                    continue
                im = Image.open(io.BytesIO(z2.read("word/media/" + relmap[rid])))
                src = im.size[1] / float(im.size[0])
                disp = int(cy) / float(cx)
                if abs(src - disp) > 0.02:
                    bad.append("%d(%.2f→%.2f)" % (i, src, disp))
            if bad:
                info["problems"].append("图片变形：第 %s 张" % "、".join(bad[:4]))
    except Exception:
        pass

    # 同行图片等高（图片大小一致对齐）
    try:
        ext_all = [(int(a), int(b)) for a, b in
                   re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc)]
        tbls_all = re.findall(r"<w:tbl>.*?</w:tbl>", doc, re.S)[1:]
        k, gaps = 0, []
        for tb in tbls_all:
            for tr in re.findall(r"<w:tr(?: [^>]*)?>.*?</w:tr>", tb, re.S):
                c = tr.count("<w:drawing>")
                hs = [ext_all[k + i][1] / 360000 for i in range(c)]
                k += c
                if len(hs) == 2 and abs(hs[0] - hs[1]) > 0.5:
                    gaps.append("%.2f/%.2f" % (hs[0], hs[1]))
        if gaps:
            info["problems"].append("同行图片不等高：%s" % "、".join(gaps[:3]))
    except Exception:
        pass

    # 颜色 / 断行
    for c in sorted(set(re.findall(r'<w:color w:val="([0-9A-Fa-f]{6})"/>', doc))):
        if c.upper() not in ("000000", "AUTO", "FFFFFF"):
            info["problems"].append("非黑文字：%s" % c)
    miss = sum(1 for t in re.findall(r"<w:tr(?: [^>]*)?>.*?</w:tr>", doc, re.S)
               if not re.search(r"<w:cantSplit(?:\s[^>]*)?/>", t))
    if miss:
        info["problems"].append("%d 行缺 cantSplit" % miss)

    # 页边距
    m = re.search(r"<w:pgMar ([^/]*)/>", doc)
    if m:
        g = dict(re.findall(r'w:(\w+)="(-?\d+)"', m.group(1)))
        if g.get("bottom") != "1440" or g.get("left") != "1080":
            info["problems"].append("页边距非模板值（bottom=%s left=%s）" % (g.get("bottom"), g.get("left")))

    # 内容口径提示：只提示、不计入 problems（不同项目口径可能不同）
    if content:
        full = text_of(doc)
        notes = info["notes"]
        if re.search(r"均为.{0,20}教师", full):
            notes.append("含内部归属声明（均为…教师），应删")
        if re.search(r"（获奖学生[^）]*指导教师", full):
            notes.append("含（获奖学生X，指导教师Y）式括注，应删")
        if re.search(r"(完成|形成|完善).{0,14}(课程标准|数字化教学资源|整体设计|技术成果|方案设计|作品制作|现场答辩)", desc) \
           or re.search(r"成立.{0,8}(工作小组|工作专班)", desc):
            notes.append("描述有成果承诺，须核对是否有对应证据图")
        if not re.search(r"(职业技术学院|技师学院|大学|本校|我校|我院)", full):
            notes.append("全文未见学校主体表述，建议用学校全称")
        y = re.search(r"(\d{4})年度任务佐证材料", info["year_line"])
        d = re.match(r"(\d{4})年", desc)
        if y and d and y.group(1) != d.group(1):
            notes.append("抬头 %s 年度 / 描述写 %s 年，年份口径需说明" % (y.group(1), d.group(1)))
        tn = re.findall(r"(\d+)\s*(?:个|项)", info["task"])
        dn = re.findall(r"(\d+)\s*(?:个|项)", desc)
        if tn and dn and not set(tn) & set(dn):
            notes.append("任务标题数量词 %s 与描述 %s 不一致" % ("/".join(tn), "/".join(dn[:3])))
    return info


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--expect", action="append", default=[], help="任务号:材料数")
    ap.add_argument("--content", action="store_true", help="追加内容口径提示（归属声明/虚报/年份/数量词，不计入问题）")
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(args.folder, "*.docx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        print("目录下没有 docx：%s" % args.folder)
        return 1

    expect = {}
    for e in args.expect:
        k, v = e.split(":")
        expect[k] = int(v)

    total = 0
    print("%-10s %-6s %-5s %-5s %-6s %-8s %s" % ("任务", "材料", "图", "描述", "图宽", "状态", "问题"))
    for f in files:
        info = check(f, content=args.content)
        m = re.search(r"任务([\d-]+)", info["file"])
        key = m.group(1) if m else info["file"][:10]
        if key in expect and expect[key] != info["materials"]:
            info["problems"].append("材料数 %d（应为 %d）" % (info["materials"], expect[key]))
        total += len(info["problems"])
        print("%-10s %-6d %-5d %-5d %-6s %-8s %s" % (
            key, info["materials"], info["images"], info["desc_len"],
            ",".join("%.2f" % w for w in info["widths"]) or "-",
            "✓" if not info["problems"] else "✗",
            "；".join(info["problems"]) or "-"))
        if args.content:
            for nt in info.get("notes", []):
                print("      · %s" % nt)
    print("\n共 %d 份，问题 %d 处" % (len(files), total))
    return 0 if total == 0 else 0


if __name__ == "__main__":
    sys.exit(main())
