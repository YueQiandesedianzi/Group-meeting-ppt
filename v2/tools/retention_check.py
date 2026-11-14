#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""保留性检查：按 ZIP 部件比较两份 PPTX 的字节哈希与大小。

规格 §10.3 要求：引擎未触碰的部件优先检查字节哈希不变；必须修改的 XML 检查只发生
允许的语义变化。本工具给出“哪些部件原样保留、哪些被重写”的事实清单。
"""
import argparse
import hashlib
import json
import os
import sys
import zipfile

CATEGORIES = (
    ("ppt/media/", "media"),
    ("ppt/slideMasters/", "master"),
    ("ppt/slideLayouts/", "layout"),
    ("ppt/slides/", "slide"),
    ("ppt/theme/", "theme"),
    ("ppt/notesSlides/", "notes"),
    ("_rels/", "package-rels"),
)


def category(name):
    for prefix, label in CATEGORIES:
        if name.startswith(prefix):
            return label
    if name.endswith(".rels"):
        return "part-rels"
    if name.startswith("docProps/"):
        return "docProps"
    return "other"


def part_hashes(path):
    out = {}
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            data = zf.read(info.filename)
            out[info.filename] = {"sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data)}
    return out


def compare(before_path, after_path):
    before = part_hashes(before_path)
    after = part_hashes(after_path)
    names = sorted(set(before) | set(after))
    parts = []
    summary = {}
    for name in names:
        b, a = before.get(name), after.get(name)
        if b and a:
            status = "identical" if b["sha256"] == a["sha256"] else "rewritten"
        elif b:
            status = "removed"
        else:
            status = "added"
        entry = {"part": name, "category": category(name), "status": status,
                 "bytes_before": b["bytes"] if b else None, "bytes_after": a["bytes"] if a else None}
        parts.append(entry)
        summary.setdefault(entry["category"], {}).setdefault(status, 0)
        summary[entry["category"]][status] += 1
    totals = {}
    for entry in parts:
        totals[entry["status"]] = totals.get(entry["status"], 0) + 1
    return {"before": os.path.abspath(before_path), "after": os.path.abspath(after_path),
            "totals": totals, "by_category": summary, "parts": parts}


def main(argv=None):
    parser = argparse.ArgumentParser(description="按部件比较 PPTX 保留情况")
    parser.add_argument("before")
    parser.add_argument("after")
    parser.add_argument("-o", "--output")
    args = parser.parse_args(argv)
    report = compare(args.before, args.after)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        with open(args.output, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
    print(json.dumps({"totals": report["totals"], "by_category": report["by_category"]},
                     ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
