"""
東証上場維持基準 監視スクリプト v3
- JPX公式XLSから銘柄一覧を取得
- stooq.com から株価取得
- Yahoo Finance から発行済株式数・決算月を取得
- 指定月の基準日銘柄を優先表示
"""

import os, sys, json, datetime, calendar, time, urllib.request, subprocess

OUTPUT_DIR = "docs"
CRITERIA = {
    "プライム":    {"danger": 150, "warning": 333},
    "スタンダード": {"danger": 15,  "warning": 33},
    "グロース":    {"danger": 8,   "warning": 17},
}
TARGET_MONTH = int(sys.argv[1]) if len(sys.argv) > 1 else datetime.date.today().month


def fetch_stock_list():
    print("[1/4] JPX銘柄一覧を取得中...")
    import pandas as pd
    xls_path = "/tmp/data_j.xls"
    req = urllib.request.Request(
        "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls",
        headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        open(xls_path, "wb").write(resp.read())
    subprocess.run(["soffice","--headless","--convert-to","csv",xls_path,"--outdir","/tmp/"],
                   check=True, capture_output=True, timeout=60)
    df = pd.read_csv("/tmp/data_j.csv", encoding="utf-8")
    df = df[df["市場・商品区分"].str.contains("内国株式", na=False)].copy()
    df["コード"] = df["コード"].astype(str).str.strip()
    df = df[df["コード"].str.match(r"^\d{4}$")].copy()
    print(f"  → {len(df)}銘柄")
    return df


def get_price_stooq(code):
    url = f"https://stooq.com/q/d/l/?s={code}.jp&i=d"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            lines = resp.read().decode("utf-8").strip().split("\n")
        if len(lines) >= 2:
            cols = lines[-1].split(",")
            if len(cols) >= 5 and cols[4] not in ("", "N/D", "null"):
                return float(cols[4])
    except Exception:
        pass
    return None


def get_shares_and_fiscal(code):
    url = (f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{code}.T"
           f"?modules=defaultKeyStatistics%2CsummaryDetail")
    try:
        req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0","Accept":"application/json"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        r = data.get("quoteSummary", {}).get("result", [{}])[0]
        shares = r.get("defaultKeyStatistics", {}).get("sharesOutstanding", {}).get("raw")
        fiscal_fmt = r.get("summaryDetail", {}).get("fiscalYearEnd", {}).get("fmt")
        m = {"january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
             "july":7,"august":8,"september":9,"october":10,"november":11,"december":12}
        fiscal_month = m.get(str(fiscal_fmt).lower()) if fiscal_fmt else None
        return shares, fiscal_month
    except Exception:
        return None, None


def fetch_stock_data(codes):
    print("[2/4] 株価・株式数・決算月を取得中...")
    results = {}
    for i, code in enumerate(codes):
        price = get_price_stooq(code)
        shares, fiscal_month = get_shares_and_fiscal(code)
        market_cap = round(price * shares / 1e8, 1) if price and shares else None
        results[code] = {"market_cap": market_cap, "fiscal_month": fiscal_month}
        if (i+1) % 200 == 0:
            got = sum(1 for v in results.values() if v["market_cap"])
            print(f"  → {i+1}/{len(codes)}件（時価総額取得済: {got}件）")
        time.sleep(0.2)
    got = sum(1 for v in results.values() if v["market_cap"])
    print(f"  → 完了: {got}/{len(codes)}件取得")
    return results


def fetch_supervision_list():
    print("[3/4] 監理・整理ポスト情報を取得中...")
    import re
    for url in [
        "https://www.jpx.co.jp/rules-participants/rules/supervision/index.html",
        "https://www.jpx.co.jp/listing/maintenance/supervision/index.html",
    ]:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                html = resp.read().decode("utf-8", errors="replace")
            result = []
            for row in re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL):
                cells = [re.sub(r'<[^>]+>', '', c).strip()
                         for c in re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)]
                if len(cells) >= 2:
                    code = re.sub(r'\D', '', cells[0])[:4]
                    if re.match(r'^\d{4}$', code):
                        result.append({"code": code, "name": cells[1],
                                       "type": "整理" if "整理" in row else "監理",
                                       "reason": cells[2] if len(cells) > 2 else ""})
            if result:
                print(f"  → {len(result)}件取得")
                return result
        except Exception as e:
            print(f"  → 失敗: {e}")
    print("  → 空リストで続行")
    return []


def generate_html(stocks_data, supervision_list, generated_at):
    print("[4/4] HTMLを生成中...")
    today = datetime.date.today()
    tm = TARGET_MONTH
    tm_ja = f"{tm}月"
    sup_codes = {s["code"] for s in supervision_list}

    dt, wt, do_, wo = [], [], [], []
    for s in stocks_data:
        is_t = s["fiscal_month"] == tm
        if s["risk"] == "danger": (dt if is_t else do_).append(s)
        elif s["risk"] == "warning": (wt if is_t else wo).append(s)

    total = sum(len(x) for x in [dt, wt, do_, wo])

    def row(s):
        code = s["code"]
        cap = f'{s["market_cap"]:,.0f}億円' if s["market_cap"] else "取得不可"
        fi = f'{s["fiscal_month"]}月決算' if s["fiscal_month"] else "不明"
        dl = f'{s["fiscal_month"]}月{calendar.monthrange(today.year,s["fiscal_month"])[1]}日' if s["fiscal_month"] else "-"
        sup = '<span class="badge bs">監理中</span>' if code in sup_codes else ""
        rb = '<span class="badge bd">危険</span>' if s["risk"]=="danger" else '<span class="badge bw">注意</span>'
        ms = s["market_str"].replace("（内国株式）","").replace("プライム","Prime").replace("スタンダード","Std").replace("グロース","Growth")
        return f'<tr class="r{s["risk"]}"><td><a href="https://finance.yahoo.co.jp/quote/{code}.T" target="_blank">{code}</a></td><td>{s["name"]}{sup}</td><td>{ms}</td><td class="rc">{cap}</td><td>{fi}</td><td class="rd">{dl}</td><td>{rb}</td></tr>'

    def sec(title, items, sid, cls):
        if not items:
            return f'<section id="{sid}"><h2 class="{cls}">{title}</h2><p class="empty">該当銘柄なし</p></section>'
        rows = "".join(row(s) for s in items)
        return f'<section id="{sid}"><h2 class="{cls}">{title} <span class="cnt">({len(items)}銘柄)</span></h2><div class="tw"><table><thead><tr><th>コード</th><th>銘柄名</th><th>市場</th><th>時価総額</th><th>決算</th><th>基準日</th><th>判定</th></tr></thead><tbody>{rows}</tbody></table></div></section>'

    sup_html = ""
    if supervision_list:
        sr = "".join(f'<tr><td>{s["code"]}</td><td>{s["name"]}</td><td><span class="badge {"bd" if s["type"]=="整理" else "bs"}">{s["type"]}ポスト</span></td><td>{s.get("reason","")}</td></tr>' for s in supervision_list)
        sup_html = f'<section id="sup"><h2 class="csup">🔴 監理・整理ポスト指定中 <span class="cnt">({len(supervision_list)}銘柄)</span></h2><div class="tw"><table><thead><tr><th>コード</th><th>銘柄名</th><th>種別</th><th>理由</th></tr></thead><tbody>{sr}</tbody></table></div></section>'

    return f"""<!DOCTYPE html>
<html lang="ja"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>東証 上場維持基準 監視サイト</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Hiragino Kaku Gothic ProN","Noto Sans JP",sans-serif;background:#f4f5f7;color:#1a1a2e;font-size:14px;line-height:1.6}}
header{{background:#1a1a2e;color:#fff;padding:20px 24px;display:flex;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;gap:8px}}
header h1{{font-size:20px;font-weight:700}}
.meta{{font-size:12px;opacity:.7;text-align:right}}
.bar{{background:#fff;border-bottom:1px solid #e0e0e0;padding:12px 24px;display:flex;gap:24px;flex-wrap:wrap;align-items:center}}
.lbl{{font-size:12px;color:#666}}.num{{font-size:22px;font-weight:700;margin-left:4px}}
.nd{{color:#d32f2f}}.nw{{color:#f57c00}}
.notice{{background:#fff8e1;border-left:4px solid #f9a825;margin:16px 24px;padding:12px 16px;font-size:12px;color:#555}}
main{{padding:0 24px 48px;max-width:1200px;margin:0 auto}}
section{{margin-top:28px}}
h2{{font-size:16px;font-weight:700;padding:10px 14px;border-radius:6px 6px 0 0;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.cdanger{{background:#ffebee;color:#b71c1c;border:1px solid #ef9a9a}}
.cwarning{{background:#fff3e0;color:#bf360c;border:1px solid #ffcc80}}
.csup{{background:#fce4ec;color:#880e4f;border:1px solid #f48fb1}}
.cnt{{font-size:13px;font-weight:400;opacity:.8}}
.tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e0e0e0;border-top:none}}
th{{background:#f5f5f5;padding:8px 12px;text-align:left;font-size:12px;color:#555;border-bottom:1px solid #ddd;white-space:nowrap}}
td{{padding:8px 12px;border-bottom:1px solid #f0f0f0;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
.rdanger{{background:#fff5f5}}.rwarning{{background:#fffbf0}}
.rdanger:hover{{background:#ffebee}}.rwarning:hover{{background:#fff3e0}}
.rc{{text-align:right}}.rd{{font-weight:600;color:#d32f2f}}
a{{color:#1565c0;text-decoration:none}}a:hover{{text-decoration:underline}}
.badge{{display:inline-block;font-size:10px;padding:2px 6px;border-radius:3px;margin-left:4px;font-weight:600;vertical-align:middle}}
.bd{{background:#d32f2f;color:#fff}}.bw{{background:#f57c00;color:#fff}}.bs{{background:#880e4f;color:#fff}}
.empty{{padding:16px;background:#fff;border:1px solid #e0e0e0;border-top:none;color:#999;font-size:13px}}
footer{{text-align:center;padding:24px;font-size:11px;color:#999;border-top:1px solid #e0e0e0;margin-top:40px}}
</style></head><body>
<header>
<div><h1>📊 東証 上場維持基準 監視サイト</h1>
<div>2025年3月〜厳格化された上場維持基準で注意が必要な銘柄を自動抽出</div></div>
<div class="meta">最終更新: {generated_at}<br>毎日自動更新（GitHub Actions）</div>
</header>
<div class="bar">
<div><span class="lbl">{tm_ja}基準日・危険</span><span class="num nd">{len(dt)}</span><span class="lbl">銘柄</span></div>
<div><span class="lbl">{tm_ja}基準日・注意</span><span class="num nw">{len(wt)}</span><span class="lbl">銘柄</span></div>
<div><span class="lbl">全月合計</span><span class="num" style="color:#1565c0">{total}</span><span class="lbl">銘柄</span></div>
</div>
<div class="notice">⚠️ <strong>免責事項：</strong>時価総額は「会社全体」（株価×発行済株式数）の推定値です。上場維持基準の「流通株式時価総額」とは異なります（流通比率30%仮定の目安）。投資判断の根拠にしないでください。基準日は決算月の月末最終日です。</div>
<main>
{sec(f"🔴 {tm_ja}が基準日・危険ゾーン", dt, "dt", "cdanger")}
{sec(f"🟠 {tm_ja}が基準日・注意ゾーン", wt, "wt", "cwarning")}
{sup_html}
{sec("⚠️ 危険ゾーン（他の月が基準日）", do_, "do", "cdanger")}
{sec("💛 注意ゾーン（他の月が基準日）", wo, "wo", "cwarning")}
</main>
<footer>データ出典：<a href="https://www.jpx.co.jp/markets/statistics-equities/misc/01.html" target="_blank">東証上場銘柄一覧（JPX）</a>、株価：stooq.com、株式数・決算月：Yahoo Finance Japan（非公式）<br>本サイトは個人学習目的で作成されており、投資助言ではありません。</footer>
</body></html>"""


def main():
    import pandas as pd
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    generated_at = datetime.datetime.now().strftime("%Y年%-m月%-d日 %H:%M")

    df = fetch_stock_list()
    codes = df["コード"].tolist()
    cap_data = fetch_stock_data(codes)
    supervision = fetch_supervision_list()

    stocks_data = []
    for _, row_data in df.iterrows():
        code = row_data["コード"]
        market_str = row_data["市場・商品区分"]
        info = cap_data.get(code, {})
        market_cap = info.get("market_cap")
        fiscal_month = info.get("fiscal_month")

        key = ("プライム" if "プライム" in market_str
               else "スタンダード" if "スタンダード" in market_str
               else "グロース" if "グロース" in market_str else None)
        if not key or market_cap is None:
            continue

        c = CRITERIA[key]
        risk = ("danger" if market_cap < c["danger"]
                else "warning" if market_cap < c["warning"] else None)
        if not risk:
            continue

        stocks_data.append({
            "code": code, "name": row_data["銘柄名"],
            "market_str": market_str, "market_cap": market_cap,
            "fiscal_month": fiscal_month, "risk": risk,
        })

    stocks_data.sort(key=lambda x: (
        x["fiscal_month"] != TARGET_MONTH,
        x["risk"] != "danger",
        x["market_cap"] or 9999
    ))

    print(f"\n対象銘柄数: {len(stocks_data)}銘柄")
    html = generate_html(stocks_data, supervision, generated_at)
    open(os.path.join(OUTPUT_DIR, "index.html"), "w", encoding="utf-8").write(html)
    json_data = {"generated_at": generated_at, "target_month": TARGET_MONTH,
                 "stocks": stocks_data, "supervision": supervision}
    open(os.path.join(OUTPUT_DIR, "data.json"), "w", encoding="utf-8").write(
        json.dumps(json_data, ensure_ascii=False, indent=2))
    print(f"✅ 完了！")

if __name__ == "__main__":
    main()
