import io
import re
from urllib.parse import urljoin

import pdfplumber
import requests

PAGES = {
    "分紅實現率": "https://www.hangseng.com/zh-hk/hangseng-insurance/fulfilment-ratio/",
    "總支付率": "https://www.hangseng.com/zh-hk/hangseng-insurance/total-payout-ratio/",
}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
KEYS = ("pdf", "fulfil", "ratio", "payout", "實現", "分紅", "支付率", "download", "下載")
MAX_PDF = 6


def fetch(url):
    note = ""
    try:
        r = requests.get(url, headers=UA, timeout=60)
        if r.status_code == 200:
            return r.content, "普通連線"
        note = f"普通連線 {r.status_code}"
    except Exception as e:
        note = f"普通連線 {type(e).__name__}"
    try:
        from curl_cffi import requests as cr
        for imp in ("chrome", "safari"):
            r = cr.get(url, impersonate=imp, timeout=60)
            if r.status_code == 200:
                return r.content, f"curl_cffi {imp}"
            note += f"｜{imp} {r.status_code}"
    except Exception as e:
        note += f"｜curl_cffi {type(e).__name__}"
    return None, note


def strip(s):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s)).strip()


def links(html, base):
    out, seen = [], set()
    for href, inner in re.findall(r'<a\b[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html, flags=re.S | re.I):
        text = strip(inner)
        url = urljoin(base, href.strip())
        blob = (url + " " + text).lower()
        if url in seen or not any(k in blob for k in KEYS):
            continue
        seen.add(url)
        out.append((url, text))
    return out


def show_tables(html):
    tables = re.findall(r"<table\b.*?</table>", html, flags=re.S | re.I)
    print(f"HTML 表格 {len(tables)} 個")
    for i, t in enumerate(tables[:6], 1):
        rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
        print(f"--- 表格 {i}：{len(rows)} 行（前 6 行） ---")
        for row in rows[:6]:
            cells = re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", row, flags=re.S | re.I)
            print("   " + " | ".join(strip(c)[:30] for c in cells))
    return len(tables)


def show_hints(html, base):
    found = re.findall(r'["\']([^"\'\s<>]+\.(?:json|pdf|xlsx?|csv)(?:\?[^"\'\s<>]*)?)["\']', html, flags=re.I)
    found += re.findall(r'<iframe\b[^>]*src=["\']([^"\']+)["\']', html, flags=re.I)
    uniq = list(dict.fromkeys(urljoin(base, u) for u in found))
    print(f"資料檔線索（json／pdf／xlsx／iframe）{len(uniq)} 條：")
    for u in uniq[:40]:
        print(f" - {u}")


def show_text(html):
    body = re.sub(r"<(script|style)\b.*?</\1>", " ", html, flags=re.S | re.I)
    lines = [strip(x) for x in re.split(r"</(?:p|div|li|h\d|tr|td|th)>|<br\s*/?>", body, flags=re.I)]
    lines = [x for x in lines if x and ("%" in x or "實現率" in x or "支付率" in x or "保單年度" in x)]
    print(f"頁面文字中有關鍵字的行 {len(lines)} 行（前 40 行）：")
    for x in lines[:40]:
        print("   " + x[:120])


def show_pdf(url):
    data, how = fetch(url)
    if data is None:
        print(f"   ✗ 下載失敗（{how}）")
        return
    if not data.startswith(b"%PDF"):
        print(f"   － 不是 PDF（{len(data) // 1024} KB，{how}）")
        return
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        meta = (pdf.metadata or {}).get("CreationDate", "")
        full = "\n".join(p.extract_text() or "" for p in pdf.pages)
        print(f"   ✓ PDF {len(data) // 1024} KB｜{len(pdf.pages)} 頁｜建立 {meta}｜{how}")
        print("   --- 文字（前 40 行） ---")
        for line in full.splitlines()[:40]:
            print("   " + line)
        for i, page in enumerate(pdf.pages[:2], 1):
            for j, table in enumerate(page.extract_tables()[:3], 1):
                cols = max(len(r) for r in table) if table else 0
                print(f"   --- 第 {i} 頁 表格 {j}：{len(table)} 行 × {cols} 欄（前 6 行） ---")
                for row in table[:6]:
                    print("   " + " | ".join(str(c or "").replace("\n", " ") for c in row))


def main():
    for name, url in PAGES.items():
        print(f"=== {name} ===")
        print(url)
        data, how = fetch(url)
        if data is None:
            print(f"✗ 頁面讀取失敗（{how}）")
            print()
            continue
        html = data.decode("utf-8", errors="replace")
        print(f"✓ 頁面 {len(html) // 1024} KB（{how}）")
        show_tables(html)
        show_hints(html, url)
        show_text(html)
        found = links(html, url)
        print(f"相關連結 {len(found)} 條：")
        for u, t in found[:60]:
            print(f" - {t[:50]} → {u}")
        print()
        pdfs = [(u, t) for u, t in found if ".pdf" in u.lower()]
        for u, t in pdfs[:MAX_PDF]:
            print(f">> {t[:50]}")
            print(f"   {u}")
            try:
                show_pdf(u)
            except Exception as e:
                print(f"   ✗ 讀取出錯：{type(e).__name__} {e}")
            print()
        if len(pdfs) > MAX_PDF:
            print(f"（PDF 共 {len(pdfs)} 份，只試前 {MAX_PDF} 份）")
            print()
    print("探查完成（只讀，沒有上傳）")


if __name__ == "__main__":
    main()
