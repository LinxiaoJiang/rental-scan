#!/usr/bin/env python3
"""Rental Scan 客户台账（Rental Scan.xlsx）的脚本化读写工具。

规则：凡是能用脚本完成的，一律用脚本完成，
禁止手工在 Excel 界面里逐格输入。本脚本是台账的唯一编辑入口。

表结构：
  第 1-2 行：备注行（类型说明 + 来源说明）
  第 3 行：表头（日期 | 客户 | 来源 | 类型 | 需求摘要 | 私信 | 回复 | 状态）
  第 4 行起：客户数据
  来源列：单元格显示平台名（Facebook / 小红书），超链接指向对应帖子/评论的原始页面。

用法示例：
  python3 rental_log.py --file <xlsx路径> list
  python3 rental_log.py --file <xlsx路径> names
  python3 rental_log.py --file <xlsx路径> add --date 2026-11-04 --name "Jane Doe" \\
      --type A --needs "长租，预算未写" --url "https://www.facebook.com/groups/xxx/posts/yyy/" \\
      --status 待联系
  python3 rental_log.py --file <xlsx路径> update --name "Jane Doe" --dm 已 --status 已联系
  python3 rental_log.py --file <xlsx路径> restyle   # 补/刷新备注行、链接样式与列宽
"""
import argparse
import datetime
import sys
import unicodedata

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HEADERS = ["日期", "客户", "来源", "类型", "需求摘要", "私信", "回复", "状态"]
COL = {h: i + 1 for i, h in enumerate(HEADERS)}
NOTES = [
    "【类型】A = 求租帖（对方发帖找房，需求基本匹配）｜B = 别人出租帖下的求租评论｜C = 自己房源帖下的新评论",
    "【来源】点平台名（Facebook / 小红书）打开对应帖子 / 评论的原页面",
]
LINK_FONT = Font(color="0563C1", underline="single")


def find_header(ws):
    for r in range(1, min(ws.max_row, 15) + 1):
        if ws.cell(row=r, column=1).value == "日期" and ws.cell(row=r, column=2).value == "客户":
            return r
    return None


def fresh_sheet():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for i, t in enumerate(NOTES, 1):
        ws.cell(row=i, column=1, value=t)
    hr = len(NOTES) + 1
    for c, h in enumerate(HEADERS, 1):
        ws.cell(row=hr, column=c, value=h)
    return wb, ws


def load(path):
    try:
        wb = openpyxl.load_workbook(path)
    except FileNotFoundError:
        return fresh_sheet()
    ws = wb.active
    if find_header(ws) is None:
        raise SystemExit(f"台账格式不对（找不到表头行）：{path}")
    return wb, ws


def data_rows(ws, hr):
    out = []
    for r in range(hr + 1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, len(HEADERS) + 1)]
        if any(v is not None for v in vals):
            out.append((r, vals))
    return out


def fmt_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v) if v is not None else ""


def disp_width(s):
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in str(s))


def autofit(ws, hr):
    for c in range(1, len(HEADERS) + 1):
        m = disp_width(HEADERS[c - 1])
        for r in range(hr + 1, ws.max_row + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                s = fmt_date(v) if c == COL["日期"] else str(v)
                m = max(m, disp_width(s))
        ws.column_dimensions[get_column_letter(c)].width = min(max(m + 2, 10), 42)


def ensure_notes(ws):
    hr = find_header(ws)
    if hr == 1:
        ws.insert_rows(1, len(NOTES))
        hr = 1 + len(NOTES)
    for i, t in enumerate(NOTES, 1):
        ws.cell(row=i, column=1, value=t)
    return hr


def set_source(cell, url, label="Facebook"):
    cell.value = label
    if url:
        cell.hyperlink = url
        cell.font = LINK_FONT


def source_display(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return cell.value or ""


def save(wb, ws, path):
    hr = find_header(ws)
    autofit(ws, hr)
    wb.save(path)


def cmd_list(ws, _args):
    hr = find_header(ws)
    print(" | ".join(HEADERS))
    for r, vals in data_rows(ws, hr):
        vals = list(vals)
        vals[0] = fmt_date(vals[0])
        vals[COL["来源"] - 1] = source_display(ws.cell(row=r, column=COL["来源"]))
        print(" | ".join("" if v is None else str(v) for v in vals))


def cmd_names(ws, _args):
    hr = find_header(ws)
    for _r, vals in data_rows(ws, hr):
        if vals[COL["客户"] - 1]:
            print(vals[COL["客户"] - 1])


def cmd_add(wb, ws, args):
    hr = ensure_notes(ws)
    for _r, vals in data_rows(ws, hr):
        if vals[COL["客户"] - 1] == args.name:
            print(f"已存在，跳过：{args.name}")
            return
    d = datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()
    r = ws.max_row + 1
    if r <= hr:
        r = hr + 1
    ws.cell(row=r, column=COL["日期"], value=d)
    ws.cell(row=r, column=COL["客户"], value=args.name)
    set_source(ws.cell(row=r, column=COL["来源"]), args.url, args.source)
    ws.cell(row=r, column=COL["类型"], value=args.type)
    ws.cell(row=r, column=COL["需求摘要"], value=args.needs or "")
    ws.cell(row=r, column=COL["私信"], value=args.dm or "未")
    ws.cell(row=r, column=COL["回复"], value=args.reply or "未")
    ws.cell(row=r, column=COL["状态"], value=args.status or "待联系")
    save(wb, ws, args.file)
    print(f"已登记：{args.name}")


def cmd_update(wb, ws, args):
    hr = find_header(ws)
    for r, vals in data_rows(ws, hr):
        if vals[COL["客户"] - 1] == args.name:
            if args.url:
                _c = ws.cell(row=r, column=COL["来源"])
                set_source(_c, args.url, args.source or _c.value or "Facebook")
            if args.dm:
                ws.cell(row=r, column=COL["私信"], value=args.dm)
            if args.reply:
                ws.cell(row=r, column=COL["回复"], value=args.reply)
            if args.status:
                ws.cell(row=r, column=COL["状态"], value=args.status)
            if args.needs is not None:
                ws.cell(row=r, column=COL["需求摘要"], value=args.needs)
            save(wb, ws, args.file)
            print(f"已更新：{args.name}")
            return
    print(f"未找到客户：{args.name}", file=sys.stderr)
    sys.exit(1)


def cmd_clear(wb, ws, args):
    hr = ensure_notes(ws)
    n = len(data_rows(ws, hr))
    if ws.max_row > hr:
        ws.delete_rows(hr + 1, ws.max_row - hr)
    save(wb, ws, args.file)
    print(f"已清空 {n} 条客户记录（备注行与表头保留）")


def cmd_restyle(wb, ws, args):
    hr = ensure_notes(ws)
    for r, _vals in data_rows(ws, hr):
        cell = ws.cell(row=r, column=COL["来源"])
        if cell.hyperlink:
            cell.font = LINK_FONT
    save(wb, ws, args.file)
    print("备注行、链接样式与列宽已刷新")


def main():
    ap = argparse.ArgumentParser(description="Rental Scan 台账工具")
    ap.add_argument("--file", required=True, help="Rental Scan.xlsx 路径")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("list")
    sub.add_parser("names")
    sub.add_parser("restyle")
    sub.add_parser("clear")
    a = sub.add_parser("add")
    a.add_argument("--date", help="YYYY-MM-DD，缺省今天")
    a.add_argument("--name", required=True)
    a.add_argument("--url", required=True, help="对应帖子/评论的原页面链接")
    a.add_argument("--source", default="Facebook", help="平台显示名：Facebook 或 小红书")
    a.add_argument("--type", required=True, choices=["A", "B", "C"])
    a.add_argument("--needs", default="")
    a.add_argument("--dm")
    a.add_argument("--reply")
    a.add_argument("--status")
    u = sub.add_parser("update")
    u.add_argument("--name", required=True)
    u.add_argument("--url")
    u.add_argument("--source")
    u.add_argument("--dm")
    u.add_argument("--reply")
    u.add_argument("--status")
    u.add_argument("--needs")
    args = ap.parse_args()

    wb, ws = load(args.file)
    if args.cmd == "list":
        cmd_list(ws, args)
    elif args.cmd == "names":
        cmd_names(ws, args)
    elif args.cmd == "add":
        cmd_add(wb, ws, args)
    elif args.cmd == "update":
        cmd_update(wb, ws, args)
    elif args.cmd == "restyle":
        cmd_restyle(wb, ws, args)
    elif args.cmd == "clear":
        cmd_clear(wb, ws, args)


if __name__ == "__main__":
    main()
