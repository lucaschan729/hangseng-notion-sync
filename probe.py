import collections
import re

import requests

PAGES = {
    "分紅實現率": "https://www.hangseng.com/zh-hk/hangseng-insurance/fulfilment-ratio/",
    "總支付率": "https://www.hangseng.com/zh-hk/hangseng-insurance/total-payout-ratio/",
}
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
S = re.S | re.I


def strip(s):
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", s).strip()


def nodes(seg):
    """回傳 [(標籤, class, 文字)]"""
    out = []
    for m in re.finditer(r"<([a-zA-Z0-9]+)\b([^>]*)>([^<]+)", seg):
        text = strip(m.group(3))
        if not text:
            continue
        cls = re.search(r'class=["\']([^"\']*)["\']', m.group(2))
        out.append((m.group(1).lower(), cls.group(1)[:50] if cls else "", text[:70]))
    return out


def cells(row):
    return [strip(c) for c in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", row, flags=S)]


def main():
    for name, url in PAGES.items():
        print(f"=== {name} ===")
        r = requests.get(url, headers=UA, timeout=60)
        if r.status_code != 200:
            print(f"✗ 讀取失敗：{r.status_code}")
            continue
        html = r.content.decode("utf-8", errors="replace")
        html = re.sub(r"<(script|style)\b.*?</\1>", " ", html, flags=S)
        html = re.sub(r"<!--.*?-->", " ", html, flags=S)

        labels = {}
        for pid, inner in re.findall(r'<a\b[^>]*href=["\'][^"\']*#(rwd-tabs[^"\']+)["\'][^>]*>(.*?)</a>',
                                     html, flags=S):
            labels[pid] = strip(inner)
        panels = [(m.start(), m.group(1))
                  for m in re.finditer(r'id=["\'](rwd-tabs[^"\']*-content)["\']', html)]
        print("分頁：", labels)
        print("分頁位置：", [(pos, labels.get(pid, pid[-12:])) for pos, pid in panels])

        others = collections.Counter()
        shown = 0
        idx = 0
        for n, m in enumerate(re.finditer(r"<table\b.*?</table>", html, flags=S), 1):
            rows = [cells(x) for x in re.findall(r"<tr\b.*?</tr>", m.group(0), flags=S)]
            rows = [x for x in rows if x]
            head = rows[0] if rows else []
            first_year = next((i for i, c in enumerate(head) if "保單年度" in c), None)
            if first_year is None or len(rows) < 2:
                others[(head[0][:30] if head else "（空）", len(rows))] += 1
                continue
            idx += 1
            panel = ""
            for pos, pid in panels:
                if pos < m.start():
                    panel = labels.get(pid, pid[-12:])
            before = nodes(html[max(0, m.start() - 5000):m.start()])
            widths = sorted({len(x) for x in rows})
            spans = len(re.findall(r"(?:row|col)span", m.group(0), flags=re.I))
            print(f"#{idx}（表格 {n}，位置 {m.start()}）｜分頁 {panel}｜{len(rows) - 1} 行"
                  f"｜每行格數 {widths}｜年度前的欄 {head[:first_year]}｜合併格 {spans}"
                  f"｜首欄 {[x[0] for x in rows[1:]]}")
            print("   前文：" + " ‖ ".join(t[2][:40] for t in before[-8:]))
            if shown < 3:
                shown += 1
                print("   --- 前 15 段文字的標籤 ---")
                for tag, cls, text in before[-15:]:
                    print(f"      <{tag} class='{cls}'> {text}")
                if shown == 1:
                    raw = re.sub(r"\s+", " ", html[max(0, m.start() - 2500):m.start()])
                    print("   --- 表格前的原始 HTML ---")
                    for i in range(0, len(raw), 300):
                        print("      " + raw[i:i + 300])
                    after = nodes(html[m.end():m.end() + 2000])
                    print("   --- 表格後 6 段文字 ---")
                    for tag, cls, text in after[:6]:
                        print(f"      <{tag} class='{cls}'> {text}")
        print(f"資料表格 {idx} 個｜其他表格（首格文字, 行數）：", dict(others))
        print()
    print("探查完成（只讀，沒有上傳）")


if __name__ == "__main__":
    main()
