import html as htmllib
import re

import requests

URL = "https://www.hangseng.com/zh-hk/hangseng-insurance/fulfilment-ratio/"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8"}
S = re.S | re.I
CJK = re.compile(r"[\u4e00-\u9fff]")
WORDS = ("計劃", "年金", "壽", "保險", "儲蓄", "高蓄")


def squash(s):
    return re.sub(r"\s+", " ", s).strip()


def blank(m):
    return " " * len(m.group(0))


def main():
    r = requests.get(URL, headers=UA, timeout=60)
    print(f"狀態 {r.status_code}")
    raw = r.content.decode("utf-8", errors="replace")
    masked = re.sub(r"<(script|style)\b.*?</\1>", blank, raw, flags=S)
    masked = re.sub(r"<!--.*?-->", blank, masked, flags=S)
    print(f"頁面 {len(raw)} 字元")

    panels = [m.start() for m in re.finditer(r'id=["\']rwd-tabs[^"\']*-content["\']', masked)]
    labels = [m.start() for m in re.finditer(r"<table\b(?:(?!</table>).)*?分紅保險產品.*?</table>", masked, flags=S)]
    print("分頁位置：", panels)
    print(f"產品標示 {len(labels)} 個，位置：", labels)
    lo = (panels[0] if panels else 0) - 8000
    hi = (labels[-1] if labels else len(raw)) + 30000
    print()

    print("=== A.「高蓄」在原始碼出現的位置 ===")
    hits = [m.start() for m in re.finditer("高蓄", raw)]
    print(f"共 {len(hits)} 次")
    for p in hits[:10]:
        print(f"[{p}] " + squash(raw[max(0, p - 200):p + 120]))
    print()

    print("=== B. 可見文字中疑似產品名稱 ===")
    seen = {}
    for m in re.finditer(r">([^<>]+)<", masked):
        t = squash(htmllib.unescape(m.group(1)))
        if lo <= m.start() <= hi and 2 <= len(t) <= 40 and any(w in t for w in WORDS):
            seen.setdefault(t, []).append(m.start())
    for t, ps in list(seen.items())[:100]:
        print(f"{len(ps)} 次｜首次 {ps[0]}｜{t}")
    print()

    print("=== C. 含中文的屬性值 ===")
    seen = {}
    for m in re.finditer(r'([\w:-]+)\s*=\s*"([^"]*)"', masked):
        name, val = m.group(1).lower(), squash(htmllib.unescape(m.group(2)))
        if lo <= m.start() <= hi and name not in ("class", "style") and CJK.search(val):
            seen.setdefault((name, val[:80]), []).append(m.start())
    for (name, val), ps in list(seen.items())[:120]:
        print(f"{len(ps)} 次｜首次 {ps[0]}｜{name} = {val}")
    print()

    print("=== D. 程式碼（script）內含中文的字串 ===")
    esc = len(re.findall(r"\\u8a08\\u5283", raw, flags=re.I))
    print(f"以 \\u 編碼的「計劃」：{esc} 次")
    seen = {}
    for sm in re.finditer(r"<script\b.*?</script>", raw, flags=S):
        for m in re.finditer(r'"([^"\\]{2,80})"', sm.group(0)):
            val = squash(m.group(1))
            if any(w in val for w in WORDS):
                seen.setdefault(val, []).append(sm.start() + m.start())
    for val, ps in list(seen.items())[:100]:
        print(f"{len(ps)} 次｜首次 {ps[0]}｜{val}")
    print()

    print("=== E.「我想查詢」之後的原始碼 ===")
    p = masked.find("我想查詢")
    if p >= 0:
        seg = re.sub(r'\s(?:class|style)="[^"]*"', "", masked[p:p + 6000])
        seg = squash(seg)[:3000]
        for i in range(0, len(seg), 300):
            print(seg[i:i + 300])
    print()

    print("=== F. 頭 3 個產品標示之前的屬性 ===")
    prev = lo
    for st in labels[:3]:
        start = max(prev, st - 6000)
        print(f"--- 標示位置 {st} ---")
        for m in re.finditer(r'([\w:-]+)\s*=\s*"([^"]*)"', masked[start:st]):
            name, val = m.group(1).lower(), squash(m.group(2))
            if name in ("id", "href", "for", "name", "value", "title", "aria-label", "aria-labelledby",
                        "aria-controls", "data-title", "data-value", "data-id", "data-target"):
                print(f"   [{start + m.start()}] {name} = {val[:100]}")
        prev = st + 200
    print()
    print("探查完成（只讀，沒有上傳）")


if __name__ == "__main__":
    main()
