#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""editor_sdk (edsdk.py) 的薄封装。

所有对 docx 的读写都通过本模块，避免每次重复拼 subprocess 调用。
用法：
    from sdk import call, nodes, tables, open_doc, save_doc
"""

import json
import os
import subprocess
import time

# 用环境变量指定 editor_sdk 技能目录；未设置时用相对目录名
SKILL_DIR = os.environ.get("WB_OFFICE_SKILL", "workbuddy-local-office-edit")
EDSDK = os.path.join(SKILL_DIR, "edsdk.py")


class SdkError(RuntimeError):
    pass


def call(tool, **kw):
    """调用 edsdk 的某个工具，返回解析后的 dict。失败抛 SdkError。"""
    proc = subprocess.run(
        ["python3", "edsdk.py", "call", tool, "--json", json.dumps(kw, ensure_ascii=False)],
        cwd=SKILL_DIR, capture_output=True, text=True,
    )
    out = (proc.stdout or "").strip()
    if proc.returncode != 0:
        raise SdkError("%s :: %s" % (tool, ((proc.stderr or out) or "")[:300]))
    if not out:
        return {}
    try:
        return json.loads(out)
    except ValueError:
        return {"raw": out}


def try_call(tool, **kw):
    """同 call，但失败只返回 None，不抛异常（用于可选的清理操作）。"""
    try:
        return call(tool, **kw)
    except SdkError:
        return None


def schema(tool):
    proc = subprocess.run(["python3", "edsdk.py", "schema", tool],
                          cwd=SKILL_DIR, capture_output=True, text=True)
    return (proc.stdout or "") + (proc.stderr or "")


def open_doc(path, wait=1.6):
    call("open_file", file_path=path)
    time.sleep(wait)


def save_doc(path):
    call("save_file", file_id=path)


def nodes(path):
    """文档结构节点；打不开时先 open 再取一次。"""
    res = call("doc_resolve_document_structure", file_id=path,
               mode="compact", limit=0, text_preview_length=200)
    ns = res.get("nodes") or []
    if not ns:
        open_doc(path)
        res = call("doc_resolve_document_structure", file_id=path,
                   mode="compact", limit=0, text_preview_length=200)
        ns = res.get("nodes") or []
    return ns


def tables(path):
    return call("doc_list_tables", file_id=path).get("tables") or []


def headings(path, level=None):
    return [n for n in nodes(path)
            if n.get("type") == "Heading"
            and (level is None or n.get("heading_level") == level)]


def paragraphs(path):
    return [n for n in nodes(path) if n.get("type") == "Paragraph"]


def table_starts(path):
    """材料表用 table_start 定位（table_id 在复制后会重复，不可靠）。"""
    return sorted(t["table_start"] for t in tables(path))


def cells_of(path, table_start):
    for t in tables(path):
        if t["table_start"] == table_start:
            return sorted(t.get("editable_cells") or [], key=lambda c: c["start"])
    return []
