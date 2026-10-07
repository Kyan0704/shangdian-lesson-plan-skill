#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_lesson_plans.py — 从解析结果(JSON) + 教案内容(JSON) 生成每班一份教案 Excel(.xlsx)。

用法:
    python generate_lesson_plans.py --plan <parsed.json> [--content <content.json>] --out <输出目录>
        [--fields "知识目标,能力目标,素养目标,教学重点,教学难点,布置作业,教学实施,课程思政"]
        [--title-len 20]
        [--template]           # 生成内容模板 JSON（供 AI/用户填写），不产出 xlsx
        [--kb <kb.json>]       # 可选：知识题库/要点难点表（按用户要求）

--plan 为 parse_teaching_plan.py 的输出。
--content 结构（班级与行序须与 --plan 一致）:
{
  "<班级名>": [
      {"topic": "原主题(用于对齐)", "unit_title": "规范化教案标题", "知识目标": "...", ...},
      ...
  ]
}
未提供 --content 时仅输出结构列（日期/周次/.../主题），内容字段留空待填；
此时若同时指定 --template，则先生成待填模板供填写后再跑本脚本。

教案 Excel 列：日期、周次、星期、节次、地点、组织形式、主题、教案标题、各内容字段。
"""

import argparse
import json
import os
import re
import sys

DEFAULT_FIELDS = ["知识目标", "能力目标", "素养目标", "教学重点", "教学难点", "布置作业", "教学实施", "课程思政"]
STRUCT_COLS = ["日期", "周次", "星期", "节次", "地点", "组织形式", "主题"]


def safe_filename(name):
    """去除文件名非法字符。"""
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(name)).strip()
    return name or "未命名班级"


def title_truncate(topic, max_len):
    """教案标题兜底：超长截断，不足则用原主题。"""
    topic = str(topic or "").strip()
    if not topic:
        return ""
    if len(topic) <= max_len:
        return topic
    # 优先在标点处截断，其次硬截
    for cut in ["：", ":", "，", ",", "、", "；", ";", "——", "-"]:
        idx = topic.find(cut)
        if 0 < idx <= max_len:
            return topic[:idx]
    return topic[:max_len]


def write_xlsx(path, headers, rows, sheet_name="教案"):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_font = Font(bold=True, size=11)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = border

    for row in rows:
        ws.append(row)

    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            cell.alignment = wrap
            cell.border = border

    # 列宽
    widths = {1: 12, 2: 6, 3: 6, 4: 10, 5: 14, 6: 10, 7: 26}
    for i, h in enumerate(headers, start=1):
        if i in widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[i]
        elif h in ("教案标题", "知识目标", "能力目标", "素养目标", "教学实施"):
            ws.column_dimensions[get_column_letter(i)].width = 30
        else:
            ws.column_dimensions[get_column_letter(i)].width = 24
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(headers)), ws.max_row)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="生成每班教案 Excel")
    ap.add_argument("--plan", required=True, help="parse_teaching_plan.py 输出的 JSON")
    ap.add_argument("--content", default=None, help="教案内容 JSON（班级→课次列表）")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--fields", default=None, help="内容字段列表，逗号分隔（默认见 DEFAULT_FIELDS）")
    ap.add_argument("--title-len", type=int, default=20, help="教案标题最大字数（默认 20）")
    ap.add_argument("--template", action="store_true", help="生成内容模板 JSON 后退出")
    ap.add_argument("--kb", default=None, help="知识题库/要点难点 JSON（可选，输出为每个班级同名的 _题库.xlsx）")
    args = ap.parse_args()

    if not os.path.isfile(args.plan):
        print("错误：--plan 文件不存在: %s" % args.plan)
        sys.exit(1)

    with open(args.plan, "r", encoding="utf-8") as f:
        plan = json.load(f)

    fields = [x.strip() for x in args.fields.split(",")] if args.fields else list(DEFAULT_FIELDS)
    classes = plan.get("classes", {})
    if not classes:
        print("错误：解析结果中没有班级数据")
        sys.exit(1)

    content = {}
    if args.content:
        if not os.path.isfile(args.content):
            print("错误：--content 文件不存在: %s" % args.content)
            sys.exit(1)
        with open(args.content, "r", encoding="utf-8") as f:
            content = json.load(f)

    if args.template:
        template = {}
        for cls, items in classes.items():
            template[cls] = [{"topic": it.get("topic", ""), "unit_title": "", **{fd: "" for fd in fields}} for it in items]
        out_t = os.path.join(args.out, "教案内容模板.json") if os.path.isdir(args.out) else args.out + "_教案内容模板.json"
        os.makedirs(os.path.dirname(out_t), exist_ok=True) if os.path.dirname(out_t) else None
        with open(out_t, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        print("内容模板已生成：%s" % os.path.abspath(out_t))
        print("请填写 unit_title 与各内容字段后，用 --content 重新运行本脚本生成教案表格。")
        sys.exit(0)

    os.makedirs(args.out, exist_ok=True)
    max_len = max(1, args.title_len)
    generated = []
    warnings = []

    for cls, items in classes.items():
        headers = STRUCT_COLS + ["教案标题"] + fields
        c_rows = content.get(cls, [])
        rows = []
        for i, it in enumerate(items):
            topic = it.get("topic", "")
            c_row = {}
            if i < len(c_rows):
                c_row = c_rows[i]
            # 主题对齐：内容行主题与计划主题不一致时给出警告
            if c_row.get("topic") and topic and c_row["topic"] != topic:
                warnings.append("%s 第 %d 条：内容主题『%s』与计划主题『%s』不一致，已按内容行取值" %
                                (cls, i + 1, c_row["topic"], topic))
            unit_title = c_row.get("unit_title") or title_truncate(topic, max_len)
            row = [
                it.get("date", ""),
                it.get("week", ""),
                it.get("weekday", ""),
                it.get("period", ""),
                it.get("location", ""),
                it.get("org_form", ""),
                topic,
                unit_title,
            ]
            row += [str(c_row.get(fd, "") or "") for fd in fields]
            rows.append(row)
        out_file = os.path.join(args.out, "%s_教案.xlsx" % safe_filename(cls))
        write_xlsx(out_file, headers, rows)
        generated.append(out_file)

        # 知识题库/要点难点（可选）
        if args.kb and os.path.isfile(args.kb):
            with open(args.kb, "r", encoding="utf-8") as f:
                kb = json.load(f)
            kb_rows = kb.get(cls)
            if kb_rows is None:
                kb_rows = kb.get("default", [])
            if kb_rows:
                kb_headers = list(kb_rows[0].keys()) if isinstance(kb_rows[0], dict) else ["要点"]
                kb_rows_out = [list(r.values()) for r in kb_rows] if isinstance(kb_rows[0], dict) else [[r] for r in kb_rows]
                kb_file = os.path.join(args.out, "%s_题库要点.xlsx" % safe_filename(cls))
                write_xlsx(kb_file, kb_headers, kb_rows_out, sheet_name="题库要点")
                generated.append(kb_file)

    print("教案表格生成完成，共 %d 个班级：" % len(generated))
    for g in generated:
        print("  - %s" % os.path.abspath(g))
    if warnings:
        print("警告：")
        for w in warnings:
            print("  - %s" % w)


if __name__ == "__main__":
    main()
