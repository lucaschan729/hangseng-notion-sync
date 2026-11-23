import bisect
import collections
import datetime
import hashlib
import html as htmllib
import json
import math
import os
import re
import sys
import time
import unicodedata

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

PAGES = {
    "分紅實現率": "https://www.hangseng.com/zh-hk/hangseng-insurance/fulfilment-ratio/",
    "總支付率": "https://www.hangseng.com/zh-hk/hangseng-insurance/total-payout-ratio/",
}
STATUS = {"不適用", "未推出", "已停售"}
EXPECT_2025 = {"分紅實現率": (29, 77, 251), "總支付率": (29, 65, 216)}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
S = re.S | re.I
PCT = re.compile(r"^(\d+(?:\.\d+)?)\s*%$")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
AUTO = "自動生成"
FORMAT = "v2"            # 改了 Excel 格式時改這個值，下次運行就會重新上傳
HINTS = {
    401: "NOTION_TOKEN 無效：請重新貼上「AIA 自動更新」的 Internal Integration Secret",
    404: "找不到資料庫：請確認資料庫已連接「AIA 自動更新」，以及 NOTION_DATABASE_ID 正確",
}
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"
NAME_COLOR = "FF1F2933"
CUR_COLOR = "FF7030A0"
SEP_COLOR = "FFE9ECEF"


class Stop(Exception):
    pass


# ---------- 讀取恒生網頁（與上一版相同） ----------
def clean(s):
    s = htmllib.unescape(re.sub(r"<[^>]+>", " ", s))
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", s)).strip()


def fetch(url):
    err = ""
    for _ in range(3):
        try:
            r = requests.get(url, headers=UA, timeout=60)
            if r.status_code == 200 and len(r.content) > 100000:
                return r.content.decode("utf-8", errors="replace")
            err = f"HTTP {r.status_code}，{len(r.content)} bytes"
        except requests.RequestException as e:
            err = type(e).__name__
        time.sleep(5)
    raise Stop(f"讀取失敗：{url}（{err}）")


def parse(name, raw):
    html = re.sub(r"<(script|style)\b.*?</\1>", " ", raw, flags=S)
    html = re.sub(r"<!--.*?-->", " ", html, flags=S)

    ym = re.search(r"(\d{4})\s*報告年度", html)
    if not ym:
        raise Stop(f"{name}：找不到報告年度")
    year = int(ym.group(1))

    labels = {}
    for pid, inner in re.findall(r'<a\b[^>]*href=["\'][^"\']*#(rwd-tabs[^"\']+)["\'][^>]*>(.*?)</a>',
                                 html, flags=S):
        labels[pid] = clean(inner)
    panels = [(m.start(), labels.get(m.group(1), ""))
              for m in re.finditer(r'id=["\'](rwd-tabs[^"\']*-content)["\']', html)]

    def panel_of(pos):
        found = ""
        for p, label in panels:
            if p < pos:
                found = label
        return found

    items, options = [], []
    for m in re.finditer(r"<div\b([^>]*)>", html):
        attrs = m.group(1)
        k = re.search(r'\bdata-key="([^"]*)"', attrs)
        if not k:
            continue
        key = clean(k.group(1))
        if re.search(r'\bdata-value="', attrs):
            options.append(key)
            continue
        c = re.search(r'\bclass="([^"]*)"', attrs)
        if c and "rwd-landing-proposition-item" in c.group(1).split():
            items.append({"pos": m.start(), "name": key, "panel": panel_of(m.start()), "tables": 0})

    if len(items) < 10:
        raise Stop(f"{name}：只找到 {len(items)} 個產品區塊，網頁結構可能已改")
    if any(not i["name"] or not i["panel"] for i in items):
        raise Stop(f"{name}：有產品缺少名稱或產品種類")
    if sorted(options) != sorted(i["name"] for i in items):
        raise Stop(f"{name}：下拉選單（{len(options)} 個）與產品區塊（{len(items)} 個）不一致")
    if len({(i["panel"], i["name"]) for i in items}) != len(items):
        raise Stop(f"{name}：有重複的產品名稱")

    starts = [i["pos"] for i in items]
    rows, years, prev_end = [], None, 0
    for m in re.finditer(r"<table\b.*?</table>", html, flags=S):
        trs = [[clean(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=S)]
               for tr in re.findall(r"<tr\b.*?</tr>", m.group(0), flags=S)]
        trs = [t for t in trs if t]
        head = trs[0] if trs else []
        if not head or head[0] != "貨幣" or not any("保單年度" in c for c in head[1:]):
            prev_end = m.end()
            continue
        idx = bisect.bisect_right(starts, m.start()) - 1
        if idx < 0:
            raise Stop(f"{name}：位置 {m.start()} 的表格不屬於任何產品")
        item = items[idx]
        titles = re.findall(r"報告年度的[^<]{1,30}", html[prev_end:m.start()])
        kind = re.sub(r"^報告年度的|之分紅實現率$", "", clean(titles[-1])).strip() if titles else ""
        if not kind:
            raise Stop(f"{name}：{item['name']} 的表格找不到標題")
        ys = []
        for c in head[1:]:
            hm = re.match(r"第\s*(\d+\+?)\s*個保單年度\D*(\d{4}(?:\s*-\s*\d{4})?)", c)
            if not hm:
                raise Stop(f"{name}：看不懂年度欄「{c}」")
            ys.append((hm.group(1), hm.group(2).replace(" ", "")))
        if years is None:
            years = ys
        elif ys != years:
            raise Stop(f"{name}：{item['name']} 的年度欄與其他表格不同")
        for tr in trs[1:]:
            if len(tr) != len(head):
                raise Stop(f"{name}：{item['name']} 有一行格數不符")
            vals = []
            for c in tr[1:]:
                pm = PCT.match(c)
                if pm:
                    vals.append(float(pm.group(1)))
                elif c in STATUS:
                    vals.append(None)
                else:
                    raise Stop(f"{name}：{item['name']} 有看不懂的內容「{c}」")
            rows.append({"panel": item["panel"], "name": item["name"], "kind": kind,
                         "ccy": tr[0], "cells": tr[1:], "vals": vals})
        item["tables"] += 1
        prev_end = m.end()

    empty = [i["name"] for i in items if not i["tables"]]
    if empty:
        raise Stop(f"{name}：以下產品沒有數據表格：{empty}")
    return {"year": year, "years": years, "rows": rows, "products": len(items)}


def stats(data):
    nums = [v for r in data["rows"] for v in r["vals"] if v is not None]
    status = {}
    for r in data["rows"]:
        for c, v in zip(r["cells"], r["vals"]):
            if v is None:
                status[c] = status.get(c, 0) + 1
    return {"products": data["products"], "rows": len(data["rows"]), "n": len(nums),
            "avg": sum(nums) / len(nums) if nums else 0, "min": min(nums, default=0),
            "max": max(nums, default=0), "low": sum(1 for x in nums if x < 80),
            "high": sum(1 for x in nums if x >= 100), "status": status}


# ---------- 保誠格式 ----------
def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid(SEP_COLOR)
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color=NAME_COLOR)
CUR_FONT = Font(bold=True, color=CUR_COLOR)
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-9 else "0.0%"


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def numbers(rows, n_fixed):
    return [v for r in rows for v in r[n_fixed:] if is_num(v)]


def to_rows(data, with_type):
    out = []
    for r in data["rows"]:
        fixed = [r["name"], r["panel"], r["ccy"]] + ([r["kind"]] if with_type else [])
        out.append(fixed + ["/" if v is None else round(v / 100, 4) for v in r["vals"]])
    return out


def year_text(n, y):
    return f"{y}或之前" if n.endswith("+") and "-" not in y else y


def make_header(years, fixed):
    return list(fixed) + [f"保单年度{n}\n{year_text(n, y)}" for n, y in years]


def build_summary(n_fixed, years, groups):
    width = n_fixed + len(years)
    head = ["所有年份平均值"] + [f"持滿{n}年平均值" for n, _ in years[1:]]
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(n_fixed, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": [year_text(n, y) for n, y in years], "year": [year_line]}


def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def cat_font(cat):
    s = str(cat)
    if "終身" in s:
        return Font(color="FF2E7D32")
    if "年金" in s:
        return Font(color="FFC55A11")
    return Font(color="FF2F5597")


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    lows = [30, 12, 10, 14]
    caps = [52, 24, 12, 44]
    widths = [fit_width([header[k]] + [row[k] for row in rows], lows[k], caps[k])
              for k in range(n_fixed)]

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 22 * (notes.count("\n") + 1)

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c == 2:
                cell.font = cat_font(v)
            elif c == 3:
                cell.font = CUR_FONT
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = n_fixed
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def add_overview(wb, year, today, items, remarks):
    """items：[(簡稱, 工作表名稱, 資料行, 固定欄數)]"""
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    def lines(r, texts):
        for t in texts:
            r += 1
            cell = ws.cell(row=r, column=1, value="・" + t)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            cell.font = NOTE_FONT
            cell.alignment = LEFT_FIT
            ws.row_dimensions[r].height = ROW_H
        return r

    banner(1, f"恒生保險分紅表現概覽（報告年度 {year}）", TITLE_FONT, DARK, 40)
    banner(2, f"資料來源：恒生保險網站｜更新日期：{today}", NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    for _, title, rows, n_fixed in items:
        nums = numbers(rows, n_fixed)
        low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
        vals = [title, len({row[0] for row in rows}), len(rows), len(nums), mean_or_slash(nums), low]
        r += 1
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
        link = ws.cell(row=r, column=1)
        link.font = LINK_FONT
        link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{title}'!A1")
        ws.cell(row=r, column=5).number_format = "0.00%"
        heat(ws.cell(row=r, column=5), vals[4], WHITE)
        ws.cell(row=r, column=6).number_format = "0.0%"
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "按產品種類的整體平均值")
    r += 1
    table_head(r, ["產品種類"] + [short for short, _, _, _ in items])
    cats = []
    for _, _, rows, _ in items:
        for row in rows:
            if row[1] not in cats:
                cats.append(row[1])
    for cat in cats:
        r += 1
        cell = ws.cell(row=r, column=1, value=cat or "（未分類）")
        cell.border = BORDER
        cell.alignment = LEFT_FIT
        cell.font = Font(bold=True, color=cat_font(cat).color)
        for k, (_, _, rows, n_fixed) in enumerate(items):
            v = mean_or_slash(numbers([row for row in rows if row[1] == cat], n_fixed))
            c2 = ws.cell(row=r, column=k + 2, value=v)
            c2.border = BORDER
            c2.alignment = CENTER_FIT
            c2.number_format = "0.00%"
            heat(c2, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方標示「不適用」、「未推出」或「已停售」")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "備註")
    r = lines(r, remarks + [f"官方來源（{n}）：{u}" for n, u in PAGES.items()])

    r += 2
    section(r, "使用提示")
    lines(r, ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
              "每個工作表第 3 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
              "向右或向下捲動時，標題列和產品名稱會固定不動。",
              "平均值摘要在每個工作表最底部。",
              "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"])


def build_excel(path, year, results, today):
    f_data, p_data = results["分紅實現率"], results["總支付率"]
    f_rows, p_rows = to_rows(f_data, True), to_rows(p_data, False)

    remarks = []
    for r in f_rows:
        top = max([v for v in r[4:] if is_num(v)] or [0])
        if top >= 3:
            remarks.append(f"{r[0]}（{r[2]}，{r[3]}）最高 {top * 100:g}%，為官方公布數字，"
                           "官方頁面設有解釋說明；整體平均值亦因此被拉高。")

    legend = ("綠色＝100% 或以上；紅色＝低於 80%；「/」＝沒有數字（官方標示不適用、未推出或已停售）。\n"
              f"資料來源：恒生保險網站（報告年度 {year}）｜更新日期：{today}")
    f_notes = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時說明的金額，定義以恒生保險網站的說明為準。"
               "100% 即與銷售時一致，高於 100% 即比銷售時多。\n" + legend)
    if remarks:
        f_notes += "\n" + "；".join(remarks)
    p_notes = ("總支付率的定義以恒生保險網站的說明為準。"
               "100% 即與銷售時的說明一致，高於 100% 即比銷售時多。\n" + legend)

    kinds = collections.Counter(r[3] for r in f_rows)
    f_groups = [("所有分紅", f_rows)]
    for k in kinds:
        sub = [r for r in f_rows if r[3] == k]
        if numbers(sub, 4):
            f_groups.append((k, sub))

    f_title, p_title = f"{year}分紅實現率", f"{year}總支付率"
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, f_title, f"恒生保險分紅實現率（報告年度 {year}）", f_notes,
              make_header(f_data["years"], ["产品名称", "产品种类", "货币", "红利类型"]), f_rows,
              build_summary(4, f_data["years"], f_groups), "分紅類型", 4, "2E75B6")
    add_sheet(wb, p_title, f"恒生保險總支付率（報告年度 {year}）", p_notes,
              make_header(p_data["years"], ["产品名称", "产品种类", "货币"]), p_rows,
              build_summary(3, p_data["years"], [("所有產品", p_rows)]), "總支付率", 3, NAVY_HEX)
    add_overview(wb, year, today,
                 [("分紅實現率", f_title, f_rows, 4), ("總支付率", p_title, p_rows, 3)], remarks)
    wb.save(path)
    print("工作表：", "、".join(wb.sheetnames))


# ---------- Notion（照保誠做法） ----------
def check(r, what):
    if r.status_code >= 300:
        hint = HINTS.get(r.status_code, "")
        raise Stop(f"{what} 失敗：{r.status_code} {r.text[:300]}" + (f"\n→ {hint}" if hint else ""))
    return r.json()


def resolve_source(raw, headers):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=headers, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=headers, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=headers, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    raise Stop(f"找不到恒生的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
               "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id, headers):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=headers, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            info = block["file"]
            cap = "".join(t.get("plain_text", "") for t in info.get("caption", []))
            out.append((block["id"], info.get("name", ""), cap))
    return out


def upload(path, year, digest, today):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    target = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not target:
        raise Stop("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    headers = {"Authorization": f"Bearer {token}", "Notion-Version": VERSION,
               "Content-Type": "application/json"}
    source = resolve_source(target, headers)
    r = requests.post(f"{API}/data_sources/{source}/query", headers=headers, json={"page_size": 100}, timeout=30)
    pages = check(r, "讀取年份資料庫").get("results", [])
    page = next((p for p in pages if str(year) in page_title(p)), None)
    if page is None:
        body = {"parent": {"type": "data_source_id", "data_source_id": source},
                "properties": {"title": {"title": [{"text": {"content": str(year)}}]}}}
        r = requests.post(f"{API}/pages", headers=headers, json=body, timeout=30)
        page = check(r, "建立年度頁面")
        print(f"已建立頁面：{year}")
    else:
        print("使用現有頁面：", page_title(page))
    page_id = page["id"]

    blocks = file_blocks(page_id, headers)
    if any(AUTO in name and digest in cap for _, name, cap in blocks):
        print("數據和格式與上次上傳相同，略過上傳")
        return
    old_auto = [bid for bid, name, _ in blocks if AUTO in name or name.startswith("HangSeng_")]

    filename = os.path.basename(path)
    r = requests.post(f"{API}/file_uploads", headers=headers,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    with open(path, "rb") as fh:
        r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                          headers={"Authorization": f"Bearer {token}", "Notion-Version": VERSION},
                          files={"file": (filename, fh.read(), XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id},
                      "caption": [{"type": "text", "text": {"content": f"更新 {today}｜{digest}"}}]}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=headers, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=headers, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 主流程 ----------
def main():
    results = {name: parse(name, fetch(url)) for name, url in PAGES.items()}
    yrs = {d["year"] for d in results.values()}
    if len(yrs) != 1:
        raise Stop(f"兩頁的報告年度不同：{yrs}")
    year = yrs.pop()
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    print(f"報告年度：{year}")

    for name, data in results.items():
        print(f"=== {name} ===")
        print("年度欄：", " ".join(f"{n}({year_text(n, y)})" for n, y in data["years"]))
        last = None
        for r in data["rows"]:
            if (r["panel"], r["name"]) != last:
                last = (r["panel"], r["name"])
                print(f"[{r['panel']}] {r['name']}")
            print(f"   {r['kind']}｜{r['ccy']}｜" + " ".join("/" if v is None else f"{v:g}%" for v in r["vals"]))

    for name, data in results.items():
        s = stats(data)
        print(f"{name}：產品 {s['products']}｜資料行 {s['rows']}｜數值 {s['n']}｜平均 {s['avg']:.2f}%"
              f"｜範圍 {s['min']:g}%–{s['max']:g}%｜低於 80%：{s['low']}｜100% 或以上：{s['high']}｜{s['status']}")
        if year == 2025:
            ok = (s["products"], s["rows"], s["n"]) == EXPECT_2025[name]
            print("   與試跑結果" + ("一致 ✓" if ok else f"不一致 ⚠（試跑為 {EXPECT_2025[name]}）"))

    path = f"恒生分紅實現率&總支付率（{year}年）{AUTO}.xlsx"
    build_excel(path, year, results, today)
    print(f"已生成 {path}（{os.path.getsize(path) // 1024} KB）")

    digest = hashlib.sha256(json.dumps(
        [FORMAT] + [[r["panel"], r["name"], r["kind"], r["ccy"], r["cells"]]
                    for d in results.values() for r in d["rows"]],
        ensure_ascii=False).encode()).hexdigest()[:12]
    if os.environ.get("UPLOAD", "").lower() == "true":
        upload(path, year, digest, today)
        print("完成")
    else:
        print("只讀試跑：沒有上傳")


if __name__ == "__main__":
    try:
        main()
    except Stop as e:
        print(f"✗ 已停止：{e}")
        sys.exit(1)
