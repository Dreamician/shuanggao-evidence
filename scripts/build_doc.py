#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按空白母版生成一份双高佐证材料。

输入：一个 JSON spec（见 references/spec-example.json）
输出：填好抬头、年度、任务标题、描述、清单、材料标题，并贴好 8cm 凭证图的 docx

流程（顺序不可调换）：
  1. 复制母版 → 目标文件
  2. XML 预处理：改文字、调整材料块数量、重塑材料表行列、清单表行数
  3. editor_sdk：倒序插入图片后保存（保存会重写表格行，故必须放在属性后处理之前）
  4. XML 后处理：文字全黑、材料标题 16pt、行加 cantSplit、图片段落居中
  5. 自检并打印报告

用法：
  python3 build_doc.py spec.json
  python3 build_doc.py spec.json --dry-run
"""

import json
import math
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sdk  # noqa: E402

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    Image = None

IMG_W_PX = 302  # 8cm @96dpi

DEFAULT_TEMPLATE = os.path.join(
    os.path.expanduser("~"),
    "双高/佐证材料-XXX专业群-任务模板.docx",
)


# ---------------------------------------------------------------- 基础工具

def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def text_of(xml):
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def iter_blocks(body_inner):
    """把 body 内部切成顶层块：[(kind, xml)]，kind ∈ p/tbl/other。"""
    out = []
    i, n = 0, len(body_inner)
    while i < n:
        if body_inner.startswith("<w:tbl>", i) or body_inner.startswith("<w:tbl ", i):
            end = match_close(body_inner, i, "w:tbl")
            out.append(("tbl", body_inner[i:end]))
            i = end
        elif body_inner.startswith("<w:p>", i) or body_inner.startswith("<w:p ", i):
            end = body_inner.find("</w:p>", i)
            end = n if end < 0 else end + len("</w:p>")
            out.append(("p", body_inner[i:end]))
            i = end
        else:
            j = i + 1
            while j < n and not (body_inner.startswith("<w:tbl", j) or body_inner.startswith("<w:p", j)):
                j += 1
            out.append(("other", body_inner[i:j]))
            i = j
    return out


def match_close(xml, start, tag):
    """从 start 处的 <tag> 找配对的 </tag>，返回结束位置（含）。

    必须带边界匹配，否则 <w:tbl 会误命中 <w:tblPr> / <w:tblGrid>。
    """
    pat = re.compile(r"<(/?)%s(?=[\s>/])[^>]*>" % re.escape(tag))
    depth = 0
    for m in pat.finditer(xml, start):
        if m.group(1):
            depth -= 1
            if depth <= 0:
                return m.end()
        else:
            if m.group(0).endswith("/>"):
                continue
            depth += 1
    return len(xml)


def set_para_text(p_xml, text):
    """替换段落文字，保留首个 run 的 rPr（字体/字号/加粗）。"""
    runs = list(re.finditer(r"<w:r(?: [^>]*)?>.*?</w:r>", p_xml, re.S))
    if not runs:
        head = re.match(r"<w:p(?: [^>]*)?>", p_xml).group(0)
        new_run = '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % esc(text)
        ppr = re.search(r"<w:pPr>.*?</w:pPr>", p_xml, re.S)
        if ppr:
            pos = ppr.end()
            return p_xml[:pos] + new_run + p_xml[pos:]
        return head + new_run + p_xml[len(head):]
    rpr = re.search(r"<w:rPr>.*?</w:rPr>", runs[0].group(0), re.S)
    new_run = "<w:r>" + (rpr.group(0) if rpr else "") + '<w:t xml:space="preserve">%s</w:t></w:r>' % esc(text)
    # 从后往前替换：先删多余 run，再换首个 run
    out = p_xml
    for m in reversed(runs[1:]):
        if "<w:t" in m.group(0):
            out = out[:m.start()] + out[m.end():]
    r0 = re.search(r"<w:r(?: [^>]*)?>.*?</w:r>", out, re.S)
    if r0:
        out = out[:r0.start()] + new_run + out[r0.end():]
    return out


def first_para(container):
    m = re.search(r"<w:p(?: [^>]*)?>.*?</w:p>", container, re.S)
    return m


def set_cell_text(tc_xml, text):
    m = first_para(tc_xml)
    if not m:
        return tc_xml
    return tc_xml[:m.start()] + set_para_text(m.group(0), text) + tc_xml[m.end():]


# ---------------------------------------------------------------- 表格构造

def table_parts(tbl_xml):
    pr = re.search(r"<w:tblPr>.*?</w:tblPr>", tbl_xml, re.S)
    grid = re.search(r"<w:tblGrid>.*?</w:tblGrid>", tbl_xml, re.S)
    rows = re.findall(r"<w:tr(?: [^>]*)?>.*?</w:tr>", tbl_xml, re.S)
    head = tbl_xml[:tbl_xml.find(">") + 1]
    return head, (pr.group(0) if pr else ""), (grid.group(0) if grid else ""), rows


def clone_rows(row_xml, times):
    return row_xml * times


def merge_row(row_xml):
    """把两列的行合并成一格（gridSpan=2）。"""
    tcs = list(re.finditer(r"<w:tc>.*?</w:tc>", row_xml, re.S))
    if len(tcs) < 2:
        return row_xml
    first, second = tcs[0].group(0), tcs[1].group(0)
    new_first = first
    w = re.search(r'<w:tcW w:w="(\d+)" w:type="(\w+)"/>', new_first)
    if w:
        val = w.group(1)
        new_val = str(int(val) * 2) if w.group(2) == "pct" else val
        new_first = new_first.replace(w.group(0),
                                      '<w:tcW w:w="%s" w:type="%s"/>' % (new_val, w.group(2)), 1)
        # gridSpan 必须紧跟 tcW
        insert_at = new_first.find('<w:tcW w:w="%s" w:type="%s"/>' % (new_val, w.group(2)))
        insert_at += len('<w:tcW w:w="%s" w:type="%s"/>' % (new_val, w.group(2)))
        new_first = new_first[:insert_at] + '<w:gridSpan w:val="2"/>' + new_first[insert_at:]
    else:
        new_first = re.sub(r"<w:tcPr>", '<w:tcPr><w:gridSpan w:val="2"/>', new_first, count=1)
    return row_xml[:tcs[0].start()] + new_first + row_xml[tcs[1].end():]


def shape_material_table(tbl_xml, n_img):
    """按图片数重塑材料表：rows = ceil(n/2)，奇数末行合并成 1 格。"""
    head, pr, grid, rows = table_parts(tbl_xml)
    if not rows:
        return tbl_xml
    # 原型必须是"两格一行"。母版材料一的末行是已合并单元格（只有 1 个 tc），
    # 单材料任务时会拿它当原型，导致整张表退化成 1 列、图片插不全。
    proto = next((r for r in rows if len(re.findall(r"<w:tc>", r)) >= 2), None)
    if proto is None:
        proto = rows[-1]
    proto = re.sub(r'<w:gridSpan w:val="\d+"/>', "", proto)  # 去掉可能残留的跨列
    rows_needed = max(1, math.ceil(n_img / 2))
    out = [proto for _ in range(rows_needed)]
    if n_img % 2 == 1:  # 最后一行只放一张图 → 合并
        out[-1] = merge_row(proto)
    return head + pr + grid + "".join(out) + "</w:tbl>"


def strip_marks(xml):
    """去掉书签标记——复制材料块时必须去掉，否则书签 id 重复。"""
    return re.sub(r"<w:bookmark(?:Start|End)[^>]*/>", "", xml)


def material_spans(blocks):
    """返回每个材料块的 (标题段下标, 表下标)。标题与表之间可能夹着 bookmarkEnd。"""
    hs = [i for i, (kind, xml) in enumerate(blocks)
          if kind == "p" and re.match(r"^材料[一二三四五六七八九十]", text_of(xml))]
    spans = []
    for h in hs:
        t = next((j for j in range(h + 1, len(blocks)) if blocks[j][0] == "tbl"), None)
        if t is not None:
            spans.append((h, t))
    return spans


def strip_cell_text(tbl_xml):
    """清空表格里所有单元格文字（保留段落与单元格属性）。"""
    def clean_tc(m):
        tc = m.group(0)
        def clean_p(pm):
            p = pm.group(0)
            for rm in reversed(list(re.finditer(r"<w:r(?: [^>]*)?>.*?</w:r>", p, re.S))):
                if "<w:t" in rm.group(0):
                    p = p[:rm.start()] + p[rm.end():]
            return p
        return re.sub(r"<w:p(?: [^>]*)?>.*?</w:p>", clean_p, tc, flags=re.S)
    return re.sub(r"<w:tc>.*?</w:tc>", clean_tc, tbl_xml, flags=re.S)


def shape_list_table(tbl_xml, n_items):
    """清单表：表头 + n_items 行。多删少补。"""
    head, pr, grid, rows = table_parts(tbl_xml)
    if not rows:
        return tbl_xml
    header, data = rows[0], rows[1:]
    while len(data) > n_items:
        data.pop()
    while len(data) < n_items:
        data.append(data[-1] if data else header)
    return head + pr + grid + header + "".join(data) + "</w:tbl>"


# ---------------------------------------------------------------- XML 后处理

def ensure_cant_split(xml):
    count = [0]

    def rep(m):
        tr = m.group(0)
        if "<w:cantSplit/>" in tr:
            return tr
        count[0] += 1
        head_m = re.match(r"<w:tr(?: [^>]*)?>", tr)
        if not head_m:
            return tr
        head = head_m.group(0)
        rest = tr[len(head):]
        trpr = re.match(r"<w:trPr>.*?</w:trPr>", rest, re.S)
        if trpr:
            new_trpr = "<w:trPr><w:cantSplit/>" + trpr.group(0)[len("<w:trPr>"):]
            return head + new_trpr + rest[trpr.end():]
        return head + "<w:trPr><w:cantSplit/></w:trPr>" + rest

    out = re.sub(r"<w:tr(?: [^>]*)?>.*?</w:tr>", rep, xml, flags=re.S)
    return out, count[0]


def ensure_center_images(xml):
    count = [0]

    def rep(m):
        p = m.group(0)
        if "<wp:inline" not in p and "<w:drawing>" not in p:
            return p
        count[0] += 1
        if re.search(r"<w:jc w:val=", p):
            return re.sub(r'<w:jc w:val="\w+"/>', '<w:jc w:val="center"/>', p, count=1)
        ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
        if ppr:
            body = ppr.group(0)
            rpr = re.search(r"<w:rPr>", body)
            if rpr:
                new_body = body[:rpr.start()] + '<w:jc w:val="center"/>' + body[rpr.start():]
            else:
                new_body = body[:-len("</w:pPr>")] + '<w:jc w:val="center"/>' + "</w:pPr>"
            return p[:ppr.start()] + new_body + p[ppr.end():]
        head = re.match(r"<w:p(?: [^>]*)?>", p).group(0)
        return p[:len(head)] + '<w:pPr><w:jc w:val="center"/></w:pPr>' + p[len(head):]

    out = re.sub(r"<w:p(?: [^>]*)?>.*?</w:p>", rep, xml, flags=re.S)
    return out, count[0]


def force_black(xml):
    def rep(m):
        val = m.group(1).upper()
        if val in ("000000", "AUTO", "FFFFFF"):
            return m.group(0)
        return '<w:color w:val="000000"/>'
    return re.sub(r'<w:color w:val="([0-9A-Fa-f]{6})"/>', rep, xml)


def bump_heading2(xml, styles_xml):
    """材料标题统一三号(16pt)：改 heading 2 样式 + 段内显式 15pt。"""
    if styles_xml:
        def fix_style(m):
            body = m.group(0)
            if "heading 2" not in body and 'w:styleId="4"' not in body:
                return body
            return body.replace('<w:sz w:val="30"/>', '<w:sz w:val="32"/>').replace(
                '<w:szCs w:val="30"/>', '<w:szCs w:val="32"/>')
        styles_xml = re.sub(r"<w:style [^>]*>.*?</w:style>", fix_style, styles_xml, flags=re.S)

    def fix_p(m):
        p = m.group(0)
        t = text_of(p)
        if re.match(r"^材料[一二三四五六七八九十]", t):
            p = p.replace('<w:sz w:val="30"/>', '<w:sz w:val="32"/>').replace(
                '<w:szCs w:val="30"/>', '<w:szCs w:val="32"/>')
        return p

    return re.sub(r"<w:p(?: [^>]*)?>.*?</w:p>", fix_p, xml, flags=re.S), styles_xml


# ---------------------------------------------------------------- 读写 zip

def read_parts(path):
    with zipfile.ZipFile(path) as z:
        doc = z.read("word/document.xml").decode("utf-8")
        styles = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in z.namelist() else None
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    return doc, styles, items


def write_parts(path, items, doc, styles):
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in items:
            if info.filename == "word/document.xml":
                data = doc.encode("utf-8")
            elif info.filename == "word/styles.xml" and styles is not None:
                data = styles.encode("utf-8")
            out.writestr(info, data)
    shutil.move(tmp, path)


# ---------------------------------------------------------------- 主流程

def img_size(path):
    """返回显示用 (w,h) 像素，宽度固定 8cm，等比。"""
    if Image is None:
        return IMG_W_PX, int(IMG_W_PX * 1.42)
    with Image.open(path) as im:
        w, h = im.size
    return IMG_W_PX, max(1, int(round(IMG_W_PX * h / float(w))))


def prepare_xml(doc, spec):
    """第 2 步：文字 + 结构。返回新的 document.xml。"""
    body_start = doc.find("<w:body>") + len("<w:body>")
    body_end = doc.rfind("</w:body>")
    inner = doc[body_start:body_end]
    blocks = iter_blocks(inner)

    # --- 2.1 四段固定文字
    targets = [
        (lambda t: t.strip() == "XXX专业群", spec["group"]),
        (lambda t: t.strip() == "年度任务佐证材料", "%s年度任务佐证材料" % spec["year"]),
        (lambda t: t.strip().startswith("任务"), spec["task"]),
        (lambda t: t.strip().startswith("写一段话描述"), spec["desc"]),
    ]
    for test, new in targets:
        for i, (kind, xml) in enumerate(blocks):
            if kind == "p" and test(text_of(xml)):
                blocks[i] = ("p", set_para_text(xml, new))
                break

    # --- 2.2 材料块数量：多了删末尾，少了复制末尾（去掉书签）
    need = len(spec["materials"])
    while True:
        spans = material_spans(blocks)
        if len(spans) <= need:
            break
        h, t = spans[-1]
        del blocks[t]   # 先删表
        del blocks[h]   # 再删标题
    while len(material_spans(blocks)) < need:
        spans = material_spans(blocks)
        h, t = spans[-1]
        seg = [("p", strip_marks(blocks[h][1])), ("tbl", strip_marks(blocks[t][1]))]
        blocks[t + 1:t + 1] = seg

    # --- 2.3 材料标题 + 材料表塑形（以最后一个块的表为原型，最干净）
    spans = material_spans(blocks)
    proto_tbl = blocks[spans[-1][1]][1]
    for k, mat in enumerate(spec["materials"]):
        h, t = spans[k]
        blocks[h] = ("p", set_para_text(blocks[h][1], mat["title"]))
        n_img = len(mat.get("images") or [])
        blocks[t] = ("tbl", strip_cell_text(shape_material_table(proto_tbl, n_img)))

    # --- 2.4 清单表：补齐/裁剪行数并写入序号与名称
    list_i = next(i for i, (kind, _) in enumerate(blocks) if kind == "tbl")
    items = spec.get("list") or []
    lt = shape_list_table(blocks[list_i][1], len(items))
    head, pr, grid, rows = table_parts(lt)
    rebuilt = []
    for r_i, tr in enumerate(rows):
        tcs = list(re.finditer(r"<w:tc>.*?</w:tc>", tr, re.S))
        if len(tcs) < 2:
            rebuilt.append(tr)
            continue
        if r_i == 0:
            tr2 = tr[:tcs[1].start()] + set_cell_text(tcs[1].group(0), "佐证材料名称") + tr[tcs[1].end():]
        else:
            tr2 = tr[:tcs[0].start()] + set_cell_text(tcs[0].group(0), str(r_i)) + tr[tcs[0].end():]
            tcs2 = list(re.finditer(r"<w:tc>.*?</w:tc>", tr2, re.S))
            val = items[r_i - 1] if r_i - 1 < len(items) else ""
            tr2 = tr2[:tcs2[1].start()] + set_cell_text(tcs2[1].group(0), val) + tr2[tcs2[1].end():]
        rebuilt.append(tr2)
    blocks[list_i] = ("tbl", head + pr + grid + "".join(rebuilt) + "</w:tbl>")

    return doc[:body_start] + "".join(x for _, x in blocks) + doc[body_end:]


def insert_images(path, spec):
    """第 3 步：倒序插图（正序会因坐标漂移而错位）。"""
    sdk.open_doc(path)
    total = 0
    for k in range(len(spec["materials"]) - 1, -1, -1):
        mat = spec["materials"][k]
        imgs = mat.get("images") or []
        if not imgs:
            continue
        starts = sdk.table_starts(path)
        mat_starts = starts[1:]  # starts[0] 是清单表
        if k >= len(mat_starts):
            raise sdk.SdkError("材料表数量不足：需要 %d 张，实际 %d 张" % (len(spec["materials"]), len(mat_starts)))
        target = mat_starts[k]
        cells = sdk.cells_of(path, target)
        n = min(len(imgs), len(cells))
        for i in range(n - 1, -1, -1):
            w, h = img_size(imgs[i])
            sdk.call("doc_insert_image", file_id=path, idx=cells[i]["start"],
                     image_path=imgs[i], w=w, h=h)
            total += 1
        print("    材料%d：%d 张图（表@%d，%d 格）" % (k + 1, n, target, len(cells)))
    sdk.save_doc(path)
    return total


def post_process(path):
    """第 4 步：属性后处理（必须在最后一次 save 之后）。"""
    doc, styles, items = read_parts(path)
    doc = force_black(doc)
    doc, styles = bump_heading2(doc, styles)
    doc, n_split = ensure_cant_split(doc)
    doc, n_center = ensure_center_images(doc)
    write_parts(path, items, doc, styles)
    print("    后处理：禁用跨页断行 %d 行，图片居中 %d 段" % (n_split, n_center))


def audit(path):
    """第 5 步：自检。"""
    with zipfile.ZipFile(path) as z:
        ok = z.testzip() is None
        doc = z.read("word/document.xml").decode("utf-8")
        styles = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in z.namelist() else ""
    problems = []
    if not ok:
        problems.append("zip 损坏")
    for kw in ("XXXX", "写一段话描述", "竖图宽度", "横图宽度", "图片相关文字部分", "XXX专业群"):
        if kw in text_of(doc):
            problems.append("残留模板占位文字：%s" % kw)
    for c in set(re.findall(r'<w:color w:val="([0-9A-Fa-f]{6})"/>', doc)):
        if c.upper() not in ("000000", "AUTO", "FFFFFF"):
            problems.append("非黑文字：%s" % c)
    miss = sum(1 for t in re.findall(r"<w:tr(?: [^>]*)?>.*?</w:tr>", doc, re.S) if "<w:cantSplit/>" not in t)
    if miss:
        problems.append("%d 行缺 cantSplit" % miss)
    imgs = doc.count("<wp:inline")
    centered = sum(1 for p in re.findall(r"<w:p(?: [^>]*)?>.*?</w:p>", doc, re.S)
                   if "<wp:inline" in p and 'w:jc w:val="center"' in p)
    if imgs and centered != imgs:
        problems.append("图片居中 %d/%d" % (centered, imgs))
    widths = sorted({round(int(a) / 360000, 2) for a, _ in re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc)})
    if widths and any(abs(w - 8.0) > 0.06 for w in widths):
        problems.append("图片宽度非 8cm：%s" % widths)
    h2 = re.findall(r"<w:style [^>]*w:styleId=\"4\"[^>]*>.*?</w:style>", styles, re.S)
    if h2 and '<w:sz w:val="32"/>' not in h2[0]:
        problems.append("heading2 样式不是 16pt")
    return {"images": imgs, "widths": widths, "problems": problems}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    with open(sys.argv[1], encoding="utf-8") as f:
        spec = json.load(f)
    dry = "--dry-run" in sys.argv
    base = os.path.dirname(os.path.abspath(sys.argv[1]))

    def absp(p):
        p = os.path.expanduser(p)
        return p if os.path.isabs(p) else os.path.join(base, p)

    tpl = absp(spec.get("template") or DEFAULT_TEMPLATE)
    out = absp(spec["out"])
    for mat in spec["materials"]:
        mat["images"] = [absp(p) for p in (mat.get("images") or [])]
    if not os.path.exists(tpl):
        print("模板不存在：%s" % tpl)
        return 1
    missing = [p for m in spec["materials"] for p in m["images"] if not os.path.exists(p)]
    if missing:
        print("图片不存在（前 5 个）：%s" % missing[:5])
        return 1
    desc_len = len(re.sub(r"\s", "", spec["desc"]))
    if not 100 <= desc_len <= 150:
        print("⚠ 描述 %d 字，要求 100-150 字" % desc_len)
    if len(spec.get("list") or []) != len(spec["materials"]):
        print("⚠ 清单 %d 条，材料 %d 个，应对应" % (len(spec.get("list") or []), len(spec["materials"])))

    print("1) 复制母版 → %s" % out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    shutil.copyfile(tpl, out)

    print("2) 预处理文字与表格")
    doc, styles, items = read_parts(out)
    doc = prepare_xml(doc, spec)
    write_parts(out, items, doc, styles)

    if dry:
        print("dry-run 结束（未插图）")
        return 0

    print("3) 插入图片")
    n = insert_images(out, spec)

    print("4) 后处理")
    post_process(out)

    print("5) 自检")
    rep = audit(out)
    print("   图片 %d 张，宽度 %s" % (rep["images"], rep["widths"]))
    if rep["problems"]:
        for p in rep["problems"]:
            print("   ⚠ %s" % p)
    else:
        print("   ✓ 无问题")
    print("完成：%s" % out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
