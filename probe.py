import html as htmllib
import re

import requests

PAGES = {
    "分紅實現率": "https://www.hangseng.com/zh-hk/hangseng-insurance/fulfilment-ratio/",
    "總支付率": "https://www.hangseng.com/zh-hk/hangseng-insurance/total-payout-ratio/",
}
EXPECT = {"分紅實現率": (29, 34, 77), "總支付率": (29, 29, 65)}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
S = re.S | re.I
LABEL = "分紅保險產品"
SKIP = re.compile(r"^(分紅保險產品|產品|\d{4}\s*年|報告年度的.*)$")
PCT = re.compile(r"^(\d+(?:\.\d+)?)\s*%$")


def strip(s):
    s = htmllib.unescape(re.sub(r"<[^>]+>", " ", s))
    return re.sub(r"\s+", " ", s).strip()


def texts(seg):
    out = [strip(t) for t in re.findall(r">([^<>]+)<", seg)]
    return [t for t in out if t]


def cells(row):
    return [strip(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", row, flags=S)]


def year_label(c):
    m = re.match(r"第\s*(\d+\+?)\s*個保單年度\D*(\d{4}(?:\s*-\s*\d{4})?)", c)
    return f"{m.group(1)}({m.group(2).replace(' ', '')})" if m else c


def kind_of(title):
    for k in ("每年紅利", "終期紅利", "總支付率"):
        if k in title:
            return k
    return title or "（找不到標題）"


def run(name, url):
    print(f"=== {name} ===")
    r = requests.get(url, headers=UA, timeout=60)
    if r.status_code != 200:
        print(f"✗ 讀取失敗：{r.status_code}")
        return
    html = r.content.decode("utf-8", errors="replace")
    html = re.sub(r"<(script|style)\b.*?</\1>", " ", html, flags=S)
    html = re.sub(r"<!--.*?-->", " ", html, flags=S)

    labels = {}
    for pid, inner in re.findall(r'<a\b[^>]*href=["\'][^"\']*#(rwd-tabs[^"\']+)["\'][^>]*>(.*?)</a>',
                                 html, flags=S):
        labels[pid] = strip(inner)
    panels = [(m.start(), labels.get(m.group(1), "？"))
              for m in re.finditer(r'id=["\'](rwd-tabs[^"\']*-content)["\']', html)]

    def panel_of(pos):
        found = "（無分頁）"
        for p, label in panels:
            if p < pos:
                found = label
        return found

    products, out, others = [], [], []
    cur, prev_end, years, n_tables = None, 0, [], 0
    for m in re.finditer(r"<table\b.*?</table>", html, flags=S):
        st, en, raw = m.start(), m.end(), m.group(0)
        rows = [cells(x) for x in re.findall(r"<tr\b.*?</tr>", raw, flags=S)]
        rows = [x for x in rows if x]
        head = rows[0] if rows else []
        first = head[0] if head else ""
        if first == LABEL:
            inside = [t for t in texts(raw) if t != LABEL]
            before = [t for t in texts(html[prev_end:st]) if not SKIP.match(t)]
            cur = {"panel": panel_of(st), "name": inside[0] if inside else None,
                   "src": "表格內" if inside else "", "end": en, "raw": raw,
                   "before": before[-5:], "between": []}
            products.append(cur)
        elif first == "貨幣" and any("保單年度" in c for c in head):
            n_tables += 1
            titles = re.findall(r"報告年度的[^<]{1,30}", html[prev_end:st])
            kind = kind_of(strip(titles[-1]) if titles else "")
            if cur is None:
                cur = {"panel": panel_of(st), "name": "（找不到產品標示）", "src": "無",
                       "end": prev_end, "raw": "", "before": [], "between": []}
                products.append(cur)
            if cur["name"] is None:
                cur["between"] = texts(html[cur["end"]:st])
                cand = [t for t in cur["between"] if not SKIP.match(t)]
                if cand:
                    cur["name"], cur["src"] = cand[0], "表格後"
                elif cur["before"]:
                    cur["name"], cur["src"] = cur["before"][-1], "表格前"
                else:
                    cur["name"], cur["src"] = "（找不到名稱）", "無"
            if not years:
                years = [year_label(c) for c in head[1:]]
            for row in rows[1:]:
                out.append((cur, kind, row[0], row[1:], len(row) == len(head)))
        else:
            others.append((st, rows))
        prev_end = en

    print("年度欄：", " ".join(years))
    print()
    last = None
    for cur, kind, ccy, vals, ok in out:
        if cur is not last:
            last = cur
            print(f"[{cur['panel']}] {(cur['name'] or '')[:60]}（名稱來源：{cur['src']}）")
        print(f"   {kind}｜{ccy}｜{' '.join(vals)}{'' if ok else '  ⚠ 格數不符'}")
    print()

    nums, status = [], {}
    for _, _, _, vals, _ in out:
        for v in vals:
            pm = PCT.match(v)
            if pm:
                nums.append(float(pm.group(1)))
            else:
                status[v] = status.get(v, 0) + 1
    src = {}
    for p in products:
        src[p["src"]] = src.get(p["src"], 0) + 1
    e = EXPECT[name]
    print(f"產品 {len(products)}（預期 {e[0]}）｜資料表格 {n_tables}（預期 {e[1]}）"
          f"｜資料行 {len(out)}（預期 {e[2]}）")
    print("名稱來源：", src)
    if nums:
        print(f"數值 {len(nums)} 個｜平均 {sum(nums) / len(nums):.2f}%｜範圍 {min(nums):g}%–{max(nums):g}%"
              f"｜低於 80%：{sum(1 for x in nums if x < 80)}｜100% 或以上：{sum(1 for x in nums if x >= 100)}")
    print("非數值格：", status)
    print()

    print("--- 診斷：頭 2 個產品標示 ---")
    for p in products[:2]:
        print("   標示表格 HTML：" + re.sub(r"\s+", " ", p["raw"])[:1200])
        print("   標示前 5 段文字：", p["before"])
        print("   標示與資料表格之間的文字：", p["between"][:10])
    print("--- 其他表格 ---")
    for st, rows in others:
        if rows and rows[0] and rows[0][0] == "產品":
            for row in rows:
                print("   [產品表] " + " | ".join(c[:60] for c in row))
        else:
            head = " | ".join(c[:15] for c in (rows[0] if rows else [])[:5])
            print(f"   位置 {st}｜{len(rows)} 行｜{head}")
    print()


def main():
    for name, url in PAGES.items():
        try:
            run(name, url)
        except Exception as e:
            print(f"✗ {name} 出錯：{type(e).__name__} {e}")
    print("試跑完成（只讀，沒有上傳）")


if __name__ == "__main__":
    main()
