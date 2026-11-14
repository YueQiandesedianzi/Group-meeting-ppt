#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""M0 往返检查：用可用的客户端重新保存 PPTX，比较结构、身份与保留信息。

Windows 无 PowerPoint 时使用 DSH 自带的 LibreOffice Kit 作为客户端代理。
报告必须注明实际使用的客户端；LibreOffice 往返不能等同于 PowerPoint 往返验收。
"""
import argparse
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pptx_inspect as pi  # noqa: E402

DEFAULT_NODE = r"D:\DSH\resources\runtime\primary-runtime\dependencies\node\bin\node.exe"
DEFAULT_CLI = (r"D:\DSH\resources\app.asar.unpacked\dsh\node_modules"
               r"\@deepseek-ai\libreoffice-kit\lib\cli.js")


def convert(node, cli, source, target, timeout=300):
    os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
    done = subprocess.run([node, cli, "convert", "--input", source, "--output", target],
                          capture_output=True, text=True, encoding="utf-8", timeout=timeout)
    if done.returncode != 0:
        raise RuntimeError("convert 失败：%s\n%s" % (done.stdout, done.stderr))
    try:
        payload = json.loads(done.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError):
        payload = {}
    return payload


def main(argv=None):
    parser = argparse.ArgumentParser(description="客户端往返与身份保留检查")
    parser.add_argument("--input", required=True)
    parser.add_argument("--converted", help="往返后的文件路径（默认 <input>.roundtrip.pptx）")
    parser.add_argument("--report", help="JSON 报告输出路径")
    parser.add_argument("--client", default="libreoffice-kit", help="往返所用客户端名称")
    parser.add_argument("--node", default=DEFAULT_NODE)
    parser.add_argument("--cli", default=DEFAULT_CLI)
    parser.add_argument("--reuse", action="store_true", help="复用已存在的 converted 文件")
    args = parser.parse_args(argv)

    source = os.path.abspath(args.input)
    converted = args.converted or os.path.splitext(source)[0] + ".roundtrip.pptx"
    meta = {}
    if not (args.reuse and os.path.isfile(converted)):
        meta = convert(args.node, args.cli, source, converted)

    before = pi.inspect_pptx(source)
    after = pi.inspect_pptx(converted)
    diff = pi.diff_reports(before, after)
    report = {
        "client": args.client,
        "client_is_powerpoint": False,
        "source": source,
        "converted": converted,
        "source_bytes": before["bytes"],
        "converted_bytes": after["bytes"],
        "convert_meta": meta,
        "sld_id_preserved": diff["sld_id_lst_before"] == diff["sld_id_lst_after"],
        "slide_size_changed": diff["slide_size_changed"],
        "slide_count": diff["slide_count"],
        "slides": diff["slides"],
        "master_markers_before": [s.get("marker") for m in before["masters"] for s in m["shapes"] if s.get("marker")],
        "master_markers_after": [s.get("marker") for m in after["masters"] for s in m["shapes"] if s.get("marker")],
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
        with open(args.report, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
