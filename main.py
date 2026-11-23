import bisect
import datetime
import hashlib
import html as htmllib
import json
import os
import re
import sys
import time
import unicodedata

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PAGES = {
    "分紅實現率": "https://www.hangseng.com/zh-hk/hangseng-insurance/fulfilment-ratio/",
    "總支付率": "https://www.hangseng.com/zh-hk/hangseng-insurance/total-payout-ratio/",
}
PAGE_TITLE = "{year}"          # Notion 頁面標題，{year} 會換成報告年度
STATUS = {"不適用", "未推出", "已停售"}
EXPECT_2025 = {"分紅實現率": (29, 77, 251), "總支付率": (29, 65, 216)}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
S = re.S | re.I
PCT = re.compile(r"^(\d+(?:\.\d+)?)\s*%$")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NOTION = "https://api.notion.com/v1"
HINTS = {
    401: "NOTION_TOKEN 無效：請到 repo 的 Settings → Secrets 重新貼上「AIA 自動更新」的 Internal Integration Secret",
    404: "找不到資料庫：請確認資料庫已連接「AIA 自動更新」，以及 NOTION_DATABASE_ID 指向原始資料庫",
}


class Stop(Exception):
    pass


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


BOLD = Font(bold=True)
PURPLE = Font(bold=True, color="7030A0")
WHITE = Font(bold=True, color="FFFFFF")
HEAD = PatternFill("solid", fgColor="1F4E78")
GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(vertical="center", wrap_text=True)


def data_sheet(wb, title, data):
    ws = wb.create_sheet(title)
    head = ["產品種類", "產品名稱", "類別", "貨幣"] + [f"第{n}年\n({y})" for n, y in data["years"]]
    ws.append(head)
    for c in ws[1]:
        c.font, c.fill, c.alignment = WHITE, HEAD, CENTER
    ws.row_dimensions[1].height = 34
    for r in data["rows"]:
        ws.append([r["panel"], r["name"], r["kind"], r["ccy"]] + ["/" if v is None else v for v in r["vals"]])
        row = ws[ws.max_row]
        row[0].alignment = row[1].alignment = row[2].alignment = LEFT
        row[1].font = BOLD
        row[3].font, row[3].alignment = PURPLE, CENTER
        for c, v in zip(row[4:], r["vals"]):
            c.alignment = CENTER
            if v is None:
                continue
            c.number_format = '0"%"' if v == int(v) else '0.0"%"'
            if v >= 100:
                c.fill = GREEN
            elif v < 80:
                c.fill = RED
        ws.row_dimensions[ws.max_row].height = 20
    for i, w in enumerate([12, 36, 12, 9] + [11] * len(data["years"]), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "E2"


def build_excel(path, year, results, now):
    wb = Workbook()
    ws = wb.active
    ws.title = "概覽"
    ws.append([f"恒生保險 {year} 報告年度：分紅實現率及總支付率"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["生成時間（香港）", now])
    for name, url in PAGES.items():
        ws.append([f"官方來源：{name}", url])
    ws.append([])
    ws.append(["項目"] + list(results))
    for c in ws[ws.max_row]:
        c.font, c.fill, c.alignment = WHITE, HEAD, CENTER
    st = {n: stats(d) for n, d in results.items()}
    for label, f in [("產品數目", lambda s: s["products"]), ("資料行（產品＋類別＋貨幣）", lambda s: s["rows"]),
                     ("數值個數", lambda s: s["n"]), ("平均", lambda s: f"{s['avg']:.2f}%"),
                     ("範圍", lambda s: f"{s['min']:g}%–{s['max']:g}%"),
                     ("100% 或以上（綠底）", lambda s: s["high"]), ("低於 80%（紅底）", lambda s: s["low"]),
                     ("顯示為 / 的格", lambda s: "、".join(f"{k} {v}" for k, v in s["status"].items()))]:
        ws.append([label] + [f(s) for s in st.values()])
    ws.append([])
    ws.append(["備註", "「不適用」、「未推出」、「已停售」一律顯示 /。數值為官方按保單年度公布的平均值。"])
    big = [r for r in results["分紅實現率"]["rows"] if any(v is not None and v >= 300 for v in r["vals"])]
    for r in big:
        ws.append(["高數值", f"{r['name']}（{r['ccy']}，{r['kind']}）最高 {max(v for v in r['vals'] if v is not None):g}%，"
                             "官方頁面設有解釋說明，請參閱官方來源。"])
    ws.append([])
    ws.append(["分紅實現率低於 80% 的資料行", "最低值"])
    for c in ws[ws.max_row]:
        c.font = BOLD
    for r in results["分紅實現率"]["rows"]:
        low = [v for v in r["vals"] if v is not None and v < 80]
        if low:
            ws.append([f"{r['name']}｜{r['kind']}｜{r['ccy']}", f"{min(low):g}%"])
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 44
    for name, data in results.items():
        data_sheet(wb, f"{year}{name}", data)
    wb.save(path)
    return st


def notion(method, path, token, **kw):
    headers = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}
    r = requests.request(method, NOTION + path, headers=headers, timeout=120, **kw)
    if r.status_code >= 300:
        hint = HINTS.get(r.status_code, "")
        raise Stop(f"Notion {method} {path.split('?')[0]} 失敗：{r.status_code} {r.text[:300]}"
                   + (f"\n→ {hint}" if hint else ""))
    return r.json()


def upload(path, year, digest, now):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    dbid = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not dbid:
        raise Stop("缺少 NOTION_TOKEN 或 NOTION_DATABASE_ID")
    db = notion("GET", f"/databases/{dbid}", token)
    sources = db.get("data_sources") or []
    if not sources:
        raise Stop("資料庫沒有 data source")
    ds = sources[0]["id"]
    props = notion("GET", f"/data_sources/{ds}", token)["properties"]
    title_prop = next(k for k, v in props.items() if v["type"] == "title")
    title = PAGE_TITLE.format(year=year)
    found = notion("POST", f"/data_sources/{ds}/query", token,
                   json={"filter": {"property": title_prop, "title": {"contains": title}}})["results"]

    def title_of(p):
        return "".join(t["plain_text"] for t in p["properties"][title_prop]["title"])

    exact = [p for p in found if title_of(p) == title]
    page = (exact or found or [None])[0]
    if page is None:
        page = notion("POST", "/pages", token, json={
            "parent": {"type": "data_source_id", "data_source_id": ds},
            "properties": {title_prop: {"title": [{"text": {"content": title}}]}}})
        print(f"已新建 Notion 頁：{title}")
    else:
        print(f"使用現有 Notion 頁：{title_of(page)}")
    pid = page["id"]

    cursor = None
    while True:
        q = f"/blocks/{pid}/children?page_size=100" + (f"&start_cursor={cursor}" if cursor else "")
        res = notion("GET", q, token)
        for b in res["results"]:
            if b["type"] == "file":
                cap = "".join(t["plain_text"] for t in b["file"].get("caption", []))
                if digest in cap:
                    print("數據與上次上傳相同，略過上傳")
                    return
        if not res.get("has_more"):
            break
        cursor = res["next_cursor"]

    name = os.path.basename(path)
    fu = notion("POST", "/file_uploads", token,
                json={"mode": "single_part", "filename": name, "content_type": XLSX})
    with open(path, "rb") as fh:
        notion("POST", f"/file_uploads/{fu['id']}/send", token, files={"file": (name, fh, XLSX)})
    caption = f"恒生保險 {year} 報告年度｜更新 {now}｜{digest}"
    notion("PATCH", f"/blocks/{pid}/children", token, json={"children": [{
        "object": "block", "type": "file",
        "file": {"type": "file_upload", "file_upload": {"id": fu["id"]},
                 "caption": [{"type": "text", "text": {"content": caption}}]}}]})
    print("已上傳 Excel 到 Notion")


def main():
    results = {name: parse(name, fetch(url)) for name, url in PAGES.items()}
    yrs = {d["year"] for d in results.values()}
    if len(yrs) != 1:
        raise Stop(f"兩頁的報告年度不同：{yrs}")
    year = yrs.pop()
    now = (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=8)).strftime("%Y-%m-%d %H:%M")
    print(f"報告年度：{year}")

    for name, data in results.items():
        print(f"=== {name} ===")
        last = None
        for r in data["rows"]:
            if (r["panel"], r["name"]) != last:
                last = (r["panel"], r["name"])
                print(f"[{r['panel']}] {r['name']}")
            print(f"   {r['kind']}｜{r['ccy']}｜" + " ".join("/" if v is None else f"{v:g}%" for v in r["vals"]))

    path = f"HangSeng_{year}_fulfilment_ratios.xlsx"
    st = build_excel(path, year, results, now)
    for name, s in st.items():
        print(f"{name}：產品 {s['products']}｜資料行 {s['rows']}｜數值 {s['n']}｜平均 {s['avg']:.2f}%"
              f"｜範圍 {s['min']:g}%–{s['max']:g}%｜低於 80%：{s['low']}｜100% 或以上：{s['high']}｜{s['status']}")
        if year == 2025:
            ok = (s["products"], s["rows"], s["n"]) == EXPECT_2025[name]
            print("   與試跑結果" + ("一致 ✓" if ok else f"不一致 ⚠（試跑為 {EXPECT_2025[name]}）"))
    print(f"已生成 {path}（{os.path.getsize(path) // 1024} KB）")

    digest = hashlib.sha256(json.dumps(
        [[r["panel"], r["name"], r["kind"], r["ccy"], r["cells"]] for d in results.values() for r in d["rows"]],
        ensure_ascii=False).encode()).hexdigest()[:12]
    if os.environ.get("UPLOAD", "").lower() == "true":
        upload(path, year, digest, now)
    else:
        print("只讀試跑：沒有上傳")


if __name__ == "__main__":
    try:
        main()
    except Stop as e:
        print(f"✗ 已停止：{e}")
        sys.exit(1)
