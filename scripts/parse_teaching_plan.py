#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_teaching_plan.py — 解析授课计划 / 教学手册文件，提取课程结构清单（每班每次课一行）。

用法:
    python parse_teaching_plan.py <input_file> [--out <output.json>]

支持格式: .xlsx / .xls / .csv / .docx / .pdf

输出 JSON:
{
  "classes": {
      "<班级名>": [
          {"class_name": "...", "week": "...", "date": "...", "weekday": "...",
           "period": "...", "location": "...", "org_form": "...", "topic": "...",
           "remark": "...", "raw": {...}},
          ...
      ]
  },
  "warnings": ["..."],
  "source": "<input_file>",
  "total_rows": N
}

说明:
- 自动探测表头（关键词匹配，多行表头时取第一个匹配行）
- 处理合并单元格（openpyxl 合并区域值向下填充）与空值前向填充（班级/日期/周次等列）
- date 列若实际为纯时间（如 "13:30-15:00"）自动改判为节次
- 关键列缺失时给出明确错误与已识别列清单，退出码非 0
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import OrderedDict

# 列名关键词（按优先级匹配，先匹配更具体的关键词）
COLUMN_KEYWORDS = [
    ("class_name", ["班级", "教学班", "授课班级", "专业班级", "班别"]),
    ("week",       ["周次", "教学周", "周数", "周别"]),
    ("weekday",    ["星期", "周几", "星期几"]),
    ("date",       ["日期", "上课日期"]),
    ("period",     ["节次", "节数", "节", "授课节次"]),
    ("time",       ["上课时间", "时间"]),
    ("location",   ["地点", "教室", "授课地点", "场地", "实训室", "机房"]),
    ("org_form",   ["组织形式", "教学形式", "授课形式", "教学方式", "组织方式"]),
    ("topic",      ["教学主题", "教学任务", "主题", "任务", "课题", "单元", "教学内容", "内容"]),
    ("remark",     ["备注", "说明"]),
]

REQUIRED_KEYS = ["topic"]
# 建议保留列（缺失时警告但不中断）
SUGGESTED_KEYS = ["class_name", "week", "date", "period", "location"]


def cell_str(v):
    """统一单元格值为字符串，去除首尾空白。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def match_column(headers, used_cols):
    """在表头中匹配各列，返回 {key: col_index}。"""
    mapping = OrderedDict()
    lower_headers = [h.lower() for h in headers]
    for key, keywords in COLUMN_KEYWORDS:
        for idx, h in enumerate(lower_headers):
            if idx in used_cols:
                continue
            for kw in keywords:
                if kw.lower() in h:
                    mapping[key] = idx
                    break
            if key in mapping:
                break
    return mapping


def read_xlsx(path):
    """读取 .xlsx，返回二维列表（合并单元格值填充）。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=False)
    ws = wb.active
    merged_map = {}
    try:
        ranges = ws.merged_cells.ranges
    except Exception:
        ranges = []
    for rng in ranges:
        val = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_map[(r, c)] = val

    def get_cell(r, c):
        if (r, c) in merged_map:
            return merged_map[(r, c)]
        return ws.cell(r, c).value

    rows = []
    for r in range(1, ws.max_row + 1):
        row = [get_cell(r, c) for c in range(1, ws.max_column + 1)]
        if any(cell_str(v) for v in row):
            rows.append(row)
    wb.close()
    return rows


def read_xls(path):
    """读取 .xls，返回二维列表。"""
    import xlrd
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows = []
    for r in range(sheet.nrows):
        row = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
        if any(cell_str(v) for v in row):
            rows.append(row)
    return rows


def read_csv(path):
    """读取 .csv，自动尝试编码，返回二维列表。"""
    encodings = ["utf-8-sig", "utf-8", "gbk", "gb18030", "latin-1"]
    last_err = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                rows = list(csv.reader(f))
            if len(rows) >= 2:
                return rows
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
    raise RuntimeError("CSV 编码识别失败: %s" % (last_err or "未知错误"))


def read_docx(path):
    """读取 .docx，优先取表格，无表格时尝试文本行。"""
    try:
        import docx
    except ImportError:
        raise RuntimeError("缺少依赖 python-docx，无法解析 .docx，请先安装")
    d = docx.Document(path)
    rows = []
    if d.tables:
        for table in d.tables:
            for row in table.rows:
                vals = [c.text for c in row.cells]
                if any(cell_str(v) for v in vals):
                    rows.append(vals)
    else:
        for p in d.paragraphs:
            t = p.text.strip()
            if t:
                rows.append(re.split(r"[\t,，;；]", t))
    return rows


def read_pdf(path):
    """读取 .pdf，优先 pdfplumber 表格，退化到文本行。"""
    rows = []
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for t in tables:
                    for row in t:
                        vals = [c if c is not None else "" for c in row]
                        if any(cell_str(v) for v in vals):
                            rows.append(vals)
                if not tables:
                    txt = page.extract_text() or ""
                    for line in txt.splitlines():
                        line = line.strip()
                        if line and not line.startswith("第") is False:
                            rows.append(re.split(r"[\t]{1,}", line))
    except ImportError:
        # 退化：PyPDF2 纯文本，仅能尽力解析
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            raise RuntimeError("缺少 pdfplumber / PyPDF2，无法解析 .pdf，请先安装")
        reader = PdfReader(path)
        for page in reader.pages:
            txt = page.extract_text() or ""
            for line in txt.splitlines():
                line = line.strip()
                if line:
                    rows.append(re.split(r"[\t]{1,}", line))
    return rows


def normalize_date(v):
    """尽力把日期格式化为 YYYY-MM-DD；失败返回原字符串。"""
    s = cell_str(v)
    if not s:
        return s
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # Excel 序列日期
        try:
            from datetime import datetime, timedelta
            return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
        except Exception:
            return s
    if isinstance(v, object) and hasattr(v, "strftime"):
        try:
            return v.strftime("%Y-%m-%d")
        except Exception:
            return s
    s2 = s.replace("/", "-").replace(".", "-").replace("年", "-").replace("月", "-").replace("日", "")
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", s2)
    if m:
        try:
            from datetime import date
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return date(y, mo, d).strftime("%Y-%m-%d")
        except Exception:
            return s
    m2 = re.search(r"(\d{1,2})月(\d{1,2})日", s)
    if m2:
        return "%s-%s" % (m2.group(1).zfill(2), m2.group(2).zfill(2))
    return s


def normalize_week(v):
    s = cell_str(v)
    if not s:
        return s
    m = re.search(r"\d+", s)
    if m and re.fullmatch(r"\D*\d+\D*", s):
        return int(m.group(0))
    return s


def is_time_only(v):
    """判断值是否纯时间（如 13:30、13:30-15:00）。"""
    s = cell_str(v)
    return bool(re.fullmatch(r"[\d:：\-\s~～至]+", s)) and ":" in s


def find_header_row(rows):
    """找表头行：至少匹配 3 个已知列关键词。"""
    for i, row in enumerate(rows[:15]):
        headers = [cell_str(v) for v in row]
        if len(headers) < 3:
            continue
        hits = 0
        for _, keywords in COLUMN_KEYWORDS:
            if any(any(kw.lower() in h.lower() for kw in keywords) for h in headers):
                hits += 1
        if hits >= 3:
            return i, headers
    return None, None


def main():
    ap = argparse.ArgumentParser(description="解析授课计划/教学手册文件")
    ap.add_argument("input_file", help="授课计划文件路径（.xlsx/.xls/.csv/.docx/.pdf）")
    ap.add_argument("--out", default=None, help="输出 JSON 路径（默认与输入同目录，同名 .json）")
    args = ap.parse_args()

    path = args.input_file
    if not os.path.isfile(path):
        print("错误：文件不存在: %s" % path)
        sys.exit(1)

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".xlsx":
            rows = read_xlsx(path)
        elif ext == ".xls":
            rows = read_xls(path)
        elif ext == ".csv":
            rows = read_csv(path)
        elif ext == ".docx":
            rows = read_docx(path)
        elif ext == ".pdf":
            rows = read_pdf(path)
        else:
            print("错误：不支持的文件格式: %s（支持 .xlsx/.xls/.csv/.docx/.pdf）" % ext)
            sys.exit(1)
    except Exception as e:
        print("解析失败：%s" % e)
        sys.exit(1)

    if not rows:
        print("错误：未读取到任何表格内容")
        sys.exit(1)

    header_idx, headers = find_header_row(rows)
    if header_idx is None:
        print("错误：无法识别表头行。请确认文件包含表头（如 班级/周次/日期/星期/节次/地点/组织形式/教学主题）。前3行示例：")
        for r in rows[:3]:
            print("  ", [cell_str(v) for v in r])
        sys.exit(1)

    used_cols = set()
    col_map = match_column(headers, used_cols)
    # 优先用独立"日期"列；若匹配到 date 但内容是纯时间，改判为 period
    if "date" in col_map:
        sample_idx = col_map["date"]
        data_samples = [cell_str(r[sample_idx]) for r in rows[header_idx + 1:] if len(r) > sample_idx]
        time_vals = [s for s in data_samples if is_time_only(s)]
        if time_vals and len(time_vals) >= max(1, len(data_samples) * 0.5) and "period" not in col_map:
            col_map["period"] = sample_idx
            del col_map["date"]

    missing = [k for k in REQUIRED_KEYS if k not in col_map]
    if missing:
        print("错误：缺少关键列: %s" % ", ".join(missing))
        print("已识别列：%s" % (", ".join("%s->%s" % (k, headers[i]) for k, i in col_map.items()) or "无"))
        sys.exit(1)

    warnings = []
    for k in SUGGESTED_KEYS:
        if k not in col_map:
            warnings.append("未识别到列: %s（该列将留空）" % k)

    # 逐行提取
    records = []
    for r in rows[header_idx + 1:]:
        if len(r) < max(col_map.values()) + 1:
            r = r + [""] * (max(col_map.values()) + 1 - len(r))
        rec = OrderedDict()
        for key, idx in col_map.items():
            rec[key] = cell_str(r[idx])
        if not any(rec.values()):
            continue
        records.append(rec)

    if not records:
        print("错误：表头之后没有数据行")
        sys.exit(1)

    # 前向填充：班级/周次/日期/星期 为空时沿用上一行（合并布局常见）
    fill_keys = ["class_name", "week", "date", "weekday"]
    last = {k: "" for k in fill_keys}
    for rec in records:
        for k in fill_keys:
            if rec.get(k):
                last[k] = rec[k]
            else:
                rec[k] = last[k]

    # 规范化与结构化
    classes = OrderedDict()
    for i, rec in enumerate(records, start=1):
        topic = rec.get("topic", "")
        if not topic:
            warnings.append("第 %d 行主题为空，已跳过" % i)
            continue
        # 时间列并入节次
        period = rec.get("period", "")
        if not period and rec.get("time"):
            period = rec["time"]
        item = {
            "class_name": rec.get("class_name", ""),
            "week": normalize_week(rec.get("week", "")),
            "date": normalize_date(rec.get("date", "")),
            "weekday": rec.get("weekday", ""),
            "period": period,
            "location": rec.get("location", ""),
            "org_form": rec.get("org_form", ""),
            "topic": topic,
            "remark": rec.get("remark", ""),
        }
        item["raw"] = {k: rec[k] for k in col_map}
        cls = item["class_name"] or "未指定班级"
        classes.setdefault(cls, []).append(item)

    if "未指定班级" in classes:
        warnings.append("未识别到班级列，全部课次归入『未指定班级』，请在生成教案时确认班级名")

    out_path = args.out or (os.path.splitext(path)[0] + ".json")
    result = {
        "classes": classes,
        "warnings": warnings,
        "source": os.path.abspath(path),
        "total_rows": sum(len(v) for v in classes.values()),
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print("解析完成：共 %d 个班级、%d 条课次记录" % (len(classes), result["total_rows"]))
    for cls, items in classes.items():
        print("  - %s: %d 条" % (cls, len(items)))
    if warnings:
        print("警告：")
        for w in warnings:
            print("  - %s" % w)
    print("输出：%s" % os.path.abspath(out_path))


if __name__ == "__main__":
    main()
