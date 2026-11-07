#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""对已有佐证材料做属性收尾/修复（不重新生成文档）。

用途：
  - 材料标题被母版继承成蓝色 → 强制改黑
  - 材料标题 15pt → 三号 16pt
  - 表格行补 cantSplit（用编辑器保存过之后会丢）
  - 图片段落补居中
  - 页边距被改过 → 恢复模板值

用法：
  python3 fix_doc.py 目录/或文件 [--margins]

注意：这些改动必须放在最后一次"用编辑器保存"之后；否则再保存会覆盖掉。
"""

import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_doc as B  # noqa: E402

TEMPLATE_MARGINS = {"top": "1440", "right": "1080", "bottom": "1440",
                    "left": "1080", "header": "851", "footer": "992"}


def restore_margins(doc):
    m = re.search(r"<w:pgMar ([^/]*)/>", doc)
    if not m:
        return doc, False
    cur = dict(re.findall(r'w:(\w+)="(-?\d+)"', m.group(1)))
    if all(cur.get(k) == v for k, v in TEMPLATE_MARGINS.items()):
        return doc, False
    attrs = " ".join('w:%s="%s"' % (k, v) for k, v in TEMPLATE_MARGINS.items())
    attrs += ' w:gutter="%s"' % cur.get("gutter", "0")
    return doc.replace(m.group(0), "<w:pgMar %s/>" % attrs), True


def fix_one(path, margins=False):
    doc, styles, items = B.read_parts(path)
    before = doc
    doc = B.force_black(doc)
    doc, styles = B.bump_heading2(doc, styles)
    doc, n_split = B.ensure_cant_split(doc)
    doc, n_center = B.ensure_center_images(doc)
    n_margin = False
    if margins:
        doc, n_margin = restore_margins(doc)
    if doc == before and not n_margin:
        return {"changed": False, "split": n_split, "center": n_center}
    B.write_parts(path, items, doc, styles)
    return {"changed": True, "split": n_split, "center": n_center, "margins": n_margin}


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else "."
    margins = "--margins" in sys.argv
    files = sorted(glob.glob(os.path.join(target, "*.docx"))) if os.path.isdir(target) else [target]
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    for f in files:
        r = fix_one(f, margins)
        flag = "已修" if r["changed"] else "无需改动"
        print("%-28s %s  cantSplit+%d 居中%d%s" % (
            os.path.basename(f)[-24:], flag, r["split"], r["center"],
            "  页边距已恢复" if r.get("margins") else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
