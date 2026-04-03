"""
東証上場維持基準 監視スクリプト v4
- JPX公式XLSから銘柄一覧を取得
- yfinance から株価・発行済株式数・決算月を取得（並列処理）
- 今月・先月の基準日銘柄を優先表示
"""

import os, sys, json, datetime, calendar, re, urllib.request, subprocess, threading
from concurrent.futures import ThreadPoolExecutor, as_completed

OUTPUT_DIR = "docs"
FISCAL_CACHE_FILE = os.path.join(OUTPUT_DIR, "fiscal_cache.json")
CRITERIA = {
    "プライム":    {"danger": 150, "warning": 333},
    "スタンダード": {"danger": 15,  "warning": 33},
    "グロース":    {"danger": 8,   "warning": 17},
}

TODAY = datetime.date.today()
TARGET_MONTH = int(sys.argv[1]) if len(sys.argv) > 1 else TODAY.month
PREV_MONTH = TARGET_MONTH - 1 if TARGET_MONTH > 1 else 12


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


def load_fiscal_cache():
    """決算月キャッシュを読み込む"""
    try:
        with open(FISCAL_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_fiscal_cache(cache):
    """決算月キャッシュを保存する"""
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(FISCAL_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass

def get_fiscal_month_kabutan(code):
    """株探から決算月をスクレイピング（yfinanceの補完用）"""
    try:
        url = f"https://kabutan.jp/stock/?code={code}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8", errors="replace")
        m = re.search(r'決算[期月][^\d]*?(\d{1,2})\s*月', html)
        if m:
            month = int(m.group(1))
            if 1 <= month <= 12:
                return month
    except Exception:
        pass
    return None

def get_stock_info(code):
    """yfinance で株価・株式数・決算月を1件取得"""
    import yfinance as yf
    try:
        info = yf.Ticker(f"{code}.T").info
        price = (info.get("currentPrice")
                 or info.get("regularMarketPrice")
                 or info.get("previousClose"))
        shares = info.get("sharesOutstanding")
        # 決算月: 複数フィールドを試す
        fiscal_month = None
        for field in ("fiscalYearEnd", "lastFiscalYearEnd", "nextFiscalYearEnd"):
            ts = info.get(field)
            if ts and isinstance(ts, (int, float)) and ts > 0:
                fiscal_month = datetime.datetime.utcfromtimestamp(ts).month
                break
        return price, shares, fiscal_month
    except Exception:
        return None, None, None


def fetch_stock_data(codes):
    print("[2/4] 株価・株式数をyfinanceで並列取得中（15ワーカー）...")
    fiscal_cache = load_fiscal_cache()
    raw = {}
    lock = threading.Lock()
    done = [0]

    def fetch_yf(code):
        price, shares, fiscal_month = get_stock_info(code)
        # yfinanceで取れなければキャッシュを使う
        if fiscal_month is None and code in fiscal_cache:
            fiscal_month = fiscal_cache[code]
        market_cap = round(price * shares / 1e8, 1) if price and shares else None
        return code, market_cap, fiscal_month

    with ThreadPoolExecutor(max_workers=15) as executor:
        futures = {executor.submit(fetch_yf, code): code for code in codes}
        for future in as_completed(futures):
            code, mc, fm = future.result()
            with lock:
                raw[code] = {"market_cap": mc, "fiscal_month": fm}
                done[0] += 1
                if done[0] % 500 == 0:
                    got_mc = sum(1 for v in raw.values() if v["market_cap"])
                    got_fm = sum(1 for v in raw.values() if v["fiscal_month"])
                    print(f"  → {done[0]}/{len(codes)}件 (時価総額:{got_mc} 決算月:{got_fm})")

    # 決算月が不明な銘柄を株探で補完
    missing = [c for c in codes if raw[c]["fiscal_month"] is None]
    if missing:
        print(f"  → 株探で決算月を補完中（{len(missing)}件 / 10ワーカー）...")
        done[0] = 0
        new_cache_entries = {}

        def fetch_kabutan(code):
            return code, get_fiscal_month_kabutan(code)

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {executor.submit(fetch_kabutan, c): c for c in missing}
            for future in as_completed(futures):
                code, fm = future.result()
                with lock:
                    if fm is not None:
                        raw[code]["fiscal_month"] = fm
                        new_cache_entries[code] = fm
                    done[0] += 1
                    if done[0] % 300 == 0:
                        print(f"  → 株探 {done[0]}/{len(missing)}件")

        if new_cache_entries:
            fiscal_cache.update(new_cache_entries)
            save_fiscal_cache(fiscal_cache)
            print(f"  → キャッシュ更新: {len(new_cache_entries)}件追加")

    got_mc = sum(1 for v in raw.values() if v["market_cap"])
    got_fm = sum(1 for v in raw.values() if v["fiscal_month"])
    print(f"  → 完了: 時価総額{got_mc}件, 決算月{got_fm}件")
    return {code: raw[code] for code in codes}


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
    tm = TARGET_MONTH
    pm = PREV_MONTH
    sup_codes = {s["code"] for s in supervision_list}

    # 月別に集計 (0 = 決算月不明)
    by_month = {m: {"danger": [], "warning": []} for m in range(0, 13)}
    for s in stocks_data:
        m = s["fiscal_month"] or 0
        by_month[m][s["risk"]].append(s)

    total = len(stocks_data)
    thead = ('<thead><tr><th>コード</th><th>銘柄名</th><th>市場</th>'
             '<th>時価総額</th><th>決算</th><th>基準日</th><th>判定</th></tr></thead>')

    def stock_row(s):
        code = s["code"]
        cap = f'{s["market_cap"]:,.0f}億円' if s["market_cap"] else "取得不可"
        fi = f'{s["fiscal_month"]}月決算' if s["fiscal_month"] else "不明"
        dl = (f'{s["fiscal_month"]}月{calendar.monthrange(TODAY.year, s["fiscal_month"])[1]}日'
              if s["fiscal_month"] else "-")
        sup = '<span class="badge bs">監理中</span>' if code in sup_codes else ""
        rb = ('<span class="badge bd">危険</span>' if s["risk"] == "danger"
              else '<span class="badge bw">注意</span>')
        ms = (s["market_str"]
              .replace("（内国株式）", "")
              .replace("プライム", "Prime")
              .replace("スタンダード", "Std")
              .replace("グロース", "Growth"))
        return (f'<tr class="r{s["risk"]}">'
                f'<td><a href="https://finance.yahoo.co.jp/quote/{code}.T" target="_blank">{code}</a></td>'
                f'<td>{s["name"]}{sup}</td><td>{ms}</td>'
                f'<td class="rc">{cap}</td><td>{fi}</td><td class="rd">{dl}</td><td>{rb}</td></tr>')

    def month_panel(m):
        d = by_month[m]["danger"]
        w = by_month[m]["warning"]
        parts = []
        for items, cls, icon, label in [
            (d, "cdanger", "🔴", "危険ゾーン"),
            (w, "cwarning", "🟠", "注意ゾーン"),
        ]:
            if items:
                rows = "".join(stock_row(s) for s in items)
                parts.append(
                    f'<div class="tsec"><h3 class="{cls}">{icon} {label} '
                    f'<span class="cnt">({len(items)}銘柄)</span></h3>'
                    f'<div class="tw"><table>{thead}<tbody>{rows}</tbody></table></div></div>')
            else:
                parts.append(
                    f'<div class="tsec"><h3 class="{cls}">{icon} {label}</h3>'
                    f'<p class="empty">該当銘柄なし</p></div>')
        return "".join(parts)

    # タブボタン生成
    tab_btns = []
    tab_panels = []

    for m in range(1, 13):
        d = by_month[m]["danger"]
        w = by_month[m]["warning"]
        badges = ""
        if d: badges += f'<span class="tbadge tbd">{len(d)}</span>'
        if w: badges += f'<span class="tbadge tbw">{len(w)}</span>'
        hl = 'data-hl="cur"' if m == tm else ('data-hl="prev"' if m == pm else "")
        active = " active" if m == tm else ""
        tab_btns.append(
            f'<button class="tab{active}" onclick="showTab({m})" id="tbtn-{m}" {hl}>'
            f'{m}月{badges}</button>')
        tab_panels.append(
            f'<div class="tab-panel{active}" id="panel-{m}">{month_panel(m)}</div>')

    # 決算月不明タブ
    ud = by_month[0]["danger"]
    uw = by_month[0]["warning"]
    if ud or uw:
        badges = ""
        if ud: badges += f'<span class="tbadge tbd">{len(ud)}</span>'
        if uw: badges += f'<span class="tbadge tbw">{len(uw)}</span>'
        tab_btns.append(
            f'<button class="tab" onclick="showTab(0)" id="tbtn-0">不明{badges}</button>')
        tab_panels.append(
            f'<div class="tab-panel" id="panel-0">{month_panel(0)}</div>')

    tab_btns_html = "\n".join(tab_btns)
    tab_panels_html = "\n".join(tab_panels)

    # 監理ポストセクション
    sup_html = ""
    if supervision_list:
        sr = "".join(
            f'<tr><td>{s["code"]}</td><td>{s["name"]}</td>'
            f'<td><span class="badge {"bd" if s["type"]=="整理" else "bs"}">{s["type"]}ポスト</span></td>'
            f'<td>{s.get("reason","")}</td></tr>'
            for s in supervision_list)
        sup_html = (f'<section id="sup"><h2 class="csup">🔴 監理・整理ポスト指定中 '
                    f'<span class="cnt">({len(supervision_list)}銘柄)</span></h2>'
                    f'<div class="tw"><table><thead><tr>'
                    f'<th>コード</th><th>銘柄名</th><th>種別</th><th>理由</th>'
                    f'</tr></thead><tbody>{sr}</tbody></table></div></section>')

    tm_d = len(by_month[tm]["danger"])
    tm_w = len(by_month[tm]["warning"])
    pm_d = len(by_month[pm]["danger"])
    pm_w = len(by_month[pm]["warning"])

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
h3{{font-size:15px;font-weight:700;padding:9px 14px;border-radius:6px 6px 0 0;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
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
/* タブ */
.tab-wrap{{background:#fff;border-bottom:2px solid #ddd;padding:16px 24px 0;margin-top:28px}}
.tab-label{{font-size:12px;color:#666;margin-bottom:8px;font-weight:600}}
.tab-bar{{display:flex;flex-wrap:wrap;gap:3px}}
.tab{{padding:7px 13px;border:1px solid #ddd;border-bottom:none;background:#f5f5f5;
      cursor:pointer;border-radius:5px 5px 0 0;font-size:13px;font-weight:500;
      color:#555;position:relative;bottom:-2px;transition:background .15s}}
.tab:hover{{background:#e8e8e8}}
.tab.active{{background:#fff;border-color:#ddd;border-bottom-color:#fff;color:#1a1a2e;font-weight:700}}
.tab[data-hl="cur"]{{border-top:3px solid #d32f2f}}
.tab[data-hl="prev"]{{border-top:3px solid #f57c00}}
.tbadge{{display:inline-block;font-size:10px;padding:1px 5px;border-radius:10px;margin-left:3px;font-weight:700}}
.tbd{{background:#d32f2f;color:#fff}}.tbw{{background:#f57c00;color:#fff}}
.tab-panels{{padding:0 24px}}
.tab-panel{{display:none}}.tab-panel.active{{display:block}}
.tsec{{margin-top:20px}}
.tab-legend{{display:flex;gap:16px;padding:10px 24px;font-size:11px;color:#888;background:#fff;border-bottom:1px solid #eee}}
.tab-legend span{{display:flex;align-items:center;gap:4px}}
</style></head><body>
<header>
<div><h1>📊 東証 上場維持基準 監視サイト</h1>
<div>2025年3月〜厳格化された上場維持基準で注意が必要な銘柄を自動抽出</div></div>
<div class="meta">最終更新: {generated_at}<br>毎日自動更新（GitHub Actions）</div>
</header>
<div class="bar">
<div><span class="lbl">今月({tm}月)危険</span><span class="num nd">{tm_d}</span><span class="lbl">銘柄</span></div>
<div><span class="lbl">今月({tm}月)注意</span><span class="num nw">{tm_w}</span><span class="lbl">銘柄</span></div>
<div><span class="lbl">先月({pm}月)危険</span><span class="num nd">{pm_d}</span><span class="lbl">銘柄</span></div>
<div><span class="lbl">先月({pm}月)注意</span><span class="num nw">{pm_w}</span><span class="lbl">銘柄</span></div>
<div><span class="lbl">全月合計</span><span class="num" style="color:#1565c0">{total}</span><span class="lbl">銘柄</span></div>
</div>
<div class="notice">⚠️ <strong>免責事項：</strong>時価総額は「会社全体」（株価×発行済株式数）の推定値です。上場維持基準の「流通株式時価総額」とは異なります（流通比率30%仮定の目安）。投資判断の根拠にしないでください。基準日は決算月の月末最終日です。</div>
<main>
{sup_html}
<div class="tab-wrap">
<div class="tab-label">📅 決算月で絞り込む</div>
<div class="tab-bar">
{tab_btns_html}
</div>
</div>
<div class="tab-legend">
<span><span class="tbadge tbd">N</span> 危険ゾーン（基準値を大きく下回る）</span>
<span><span class="tbadge tbw">N</span> 注意ゾーン（基準値に近い）</span>
<span style="color:#d32f2f">■</span> 今月({tm}月)が基準日&nbsp;
<span style="color:#f57c00">■</span> 先月({pm}月)が基準日
</div>
<div class="tab-panels">
{tab_panels_html}
</div>
</main>
<footer>データ出典：<a href="https://www.jpx.co.jp/markets/statistics-equities/misc/01.html" target="_blank">東証上場銘柄一覧（JPX）</a>、株価・株式数・決算月：Yahoo Finance（yfinance）<br>本サイトは個人学習目的で作成されており、投資助言ではありません。</footer>
<script>
function showTab(m) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
  var panel = document.getElementById('panel-' + m);
  var btn = document.getElementById('tbtn-' + m);
  if (panel) panel.classList.add('active');
  if (btn) btn.classList.add('active');
}}
</script>
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
        0 if x["fiscal_month"] == TARGET_MONTH else
        1 if x["fiscal_month"] == PREV_MONTH else 2,
        0 if x["risk"] == "danger" else 1,
        x["market_cap"] or 9999
    ))

    print(f"\n対象銘柄数: {len(stocks_data)}銘柄")
    html = generate_html(stocks_data, supervision, generated_at)
    open(os.path.join(OUTPUT_DIR, "index.html"), "w", encoding="utf-8").write(html)
    json_data = {"generated_at": generated_at, "target_month": TARGET_MONTH,
                 "prev_month": PREV_MONTH, "stocks": stocks_data, "supervision": supervision}
    open(os.path.join(OUTPUT_DIR, "data.json"), "w", encoding="utf-8").write(
        json.dumps(json_data, ensure_ascii=False, indent=2))
    print("✅ 完了！")

if __name__ == "__main__":
    main()
