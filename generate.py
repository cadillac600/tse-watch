"""
東証上場維持基準 監視スクリプト v5
- JPX公式XLSから銘柄一覧を取得
- yfinance から株価・浮動株数・決算月・上場日を取得（並列）
- 流通株式時価総額（floatShares×株価）で正確判定
- 流通株式比率チェック追加
- 監理ポスト・改善期間該当銘柄を正しいURLから取得
"""

import os, sys, json, datetime, calendar, re, urllib.request, subprocess, threading
from concurrent.futures import ThreadPoolExecutor, as_completed

OUTPUT_DIR = "docs"
FISCAL_CACHE_FILE = os.path.join(OUTPUT_DIR, "fiscal_cache.json")

# 上場維持基準（流通株式時価総額ベース・単位：億円）
# float_danger = 実際の基準値（これを下回ると違反）
# float_warning = 基準値の2倍以内（近づいている）
# ratio_min = 流通株式比率の最低値（%）
CRITERIA = {
    "プライム":    {"float_danger": 100, "float_warning": 200, "ratio_min": 35},
    "スタンダード": {"float_danger": 10,  "float_warning": 20,  "ratio_min": 25},
    "グロース":    {"float_danger": 5,   "float_warning": 10,  "ratio_min": 25},
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


# ── 決算月キャッシュ ──────────────────────────────────
def load_fiscal_cache():
    try:
        with open(FISCAL_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_fiscal_cache(cache):
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(FISCAL_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass

def get_fiscal_month_kabutan(code):
    """株探から決算月をスクレイピング"""
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


# ── yfinance データ取得 ──────────────────────────────
def get_stock_info(code):
    """
    yfinance で1銘柄のデータを取得。
    Returns: (total_cap, float_cap, float_ratio, fiscal_month, listing_date)
      - total_cap  : 会社全体の時価総額 (億円)
      - float_cap  : 流通株式時価総額 (億円) ← floatShares×株価
      - float_ratio: 流通株式比率 (%) = floatShares/sharesOutstanding×100
      - fiscal_month: 決算月 (1-12 or None)
      - listing_date: 上場日 (date or None)
    """
    import yfinance as yf
    try:
        info = yf.Ticker(f"{code}.T").info
        price  = (info.get("currentPrice")
                  or info.get("regularMarketPrice")
                  or info.get("previousClose"))
        shares       = info.get("sharesOutstanding")
        float_shares = info.get("floatShares")

        total_cap  = round(price * shares / 1e8, 1) if price and shares else None
        float_cap  = round(price * float_shares / 1e8, 1) if price and float_shares else None
        float_ratio = (round(float_shares / shares * 100, 1)
                       if float_shares and shares else None)

        # 決算月
        fiscal_month = None
        for field in ("fiscalYearEnd", "lastFiscalYearEnd", "nextFiscalYearEnd"):
            ts = info.get(field)
            if ts and isinstance(ts, (int, float)) and ts > 0:
                fiscal_month = datetime.datetime.utcfromtimestamp(ts).month
                break

        # 上場日
        listing_ts   = info.get("firstTradeDateEpochUtc")
        listing_date = (datetime.date.fromtimestamp(listing_ts)
                        if listing_ts and listing_ts > 0 else None)

        return total_cap, float_cap, float_ratio, fiscal_month, listing_date
    except Exception:
        return None, None, None, None, None


def fetch_stock_data(codes):
    print("[2/4] 株価・浮動株数・決算月を並列取得中（yfinance / 15ワーカー）...")
    fiscal_cache = load_fiscal_cache()
    raw  = {}
    lock = threading.Lock()
    done = [0]

    def fetch_yf(code):
        total_cap, float_cap, float_ratio, fiscal_month, listing_date = get_stock_info(code)
        if fiscal_month is None and code in fiscal_cache:
            fiscal_month = fiscal_cache[code]
        return code, total_cap, float_cap, float_ratio, fiscal_month, listing_date

    with ThreadPoolExecutor(max_workers=15) as executor:
        futures = {executor.submit(fetch_yf, c): c for c in codes}
        for fut in as_completed(futures):
            c, tc, fc, fr, fm, ld = fut.result()
            with lock:
                raw[c] = {"total_cap": tc, "float_cap": fc, "float_ratio": fr,
                          "fiscal_month": fm, "listing_date": ld}
                done[0] += 1
                if done[0] % 500 == 0:
                    got_tc = sum(1 for v in raw.values() if v["total_cap"])
                    got_fc = sum(1 for v in raw.values() if v["float_cap"])
                    got_fm = sum(1 for v in raw.values() if v["fiscal_month"])
                    print(f"  → {done[0]}/{len(codes)}件 "
                          f"(総時価:{got_tc} 流通時価:{got_fc} 決算月:{got_fm})")

    # 決算月が不明な銘柄を株探で補完
    missing = [c for c in codes if raw[c]["fiscal_month"] is None]
    if missing:
        print(f"  → 株探で決算月を補完中（{len(missing)}件 / 10ワーカー）...")
        done[0] = 0
        new_entries = {}

        def fetch_kab(code):
            return code, get_fiscal_month_kabutan(code)

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {executor.submit(fetch_kab, c): c for c in missing}
            for fut in as_completed(futures):
                c, fm = fut.result()
                with lock:
                    if fm is not None:
                        raw[c]["fiscal_month"] = fm
                        new_entries[c] = fm
                    done[0] += 1
                    if done[0] % 300 == 0:
                        print(f"  → 株探 {done[0]}/{len(missing)}件")

        if new_entries:
            fiscal_cache.update(new_entries)
            save_fiscal_cache(fiscal_cache)
            print(f"  → キャッシュ更新: {len(new_entries)}件")

    got_fc = sum(1 for v in raw.values() if v["float_cap"])
    got_fm = sum(1 for v in raw.values() if v["fiscal_month"])
    print(f"  → 完了: 流通時価総額{got_fc}件, 決算月{got_fm}件")
    return {c: raw[c] for c in codes}


# ── 監理ポスト・改善期間 取得 ─────────────────────────
def _scrape_jpx_alert_page(url, post_type, pd_module):
    """JPXの監理/改善期間ページをスクレイピング"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "text/html,application/xhtml+xml,*/*;q=0.9",
        "Accept-Language": "ja,en-US;q=0.7",
        "Referer": "https://www.jpx.co.jp/listing/",
    }
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=15) as resp:
        html = resp.read().decode("utf-8", errors="replace")
    result = []
    # pandas.read_html 優先
    try:
        tables = pd_module.read_html(html)
        for df in tables:
            for col in df.columns:
                series = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                if series.str.match(r'^\d{4}$').sum() >= 3:
                    col_idx = list(df.columns).index(col)
                    name_col = (df.columns[col_idx + 1]
                                if col_idx + 1 < len(df.columns) else None)
                    for _, row in df.iterrows():
                        code = str(row[col]).strip().replace('.0', '')
                        if re.match(r'^\d{4}$', code):
                            name = str(row[name_col]).strip() if name_col else ""
                            result.append({"code": code, "name": name,
                                           "type": post_type, "reason": ""})
                    break
    except Exception:
        pass
    # regex フォールバック
    if not result:
        for row_html in re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL):
            cells = [re.sub(r'<[^>]+>', '', c).strip()
                     for c in re.findall(r'<td[^>]*>(.*?)</td>', row_html, re.DOTALL)]
            for i, cell in enumerate(cells[:3]):
                m = re.search(r'\b(\d{4})\b', cell)
                if m:
                    result.append({"code": m.group(1),
                                   "name": cells[i + 1] if i + 1 < len(cells) else "",
                                   "type": post_type, "reason": ""})
                    break
    return result


def fetch_supervision_list():
    import pandas as pd
    print("[3/4] 監理・整理・改善期間 情報を取得中...")
    all_result = []
    pages = [
        # 正しいURL（monitoring alerts）
        ("https://www.jpx.co.jp/listing/market-alerts/supervision/index.html",       "監理"),
        ("https://www.jpx.co.jp/listing/market-alerts/seiri/index.html",             "整理"),
        ("https://www.jpx.co.jp/listing/market-alerts/improvement-period/index.html","改善期間"),
    ]
    for url, post_type in pages:
        try:
            result = _scrape_jpx_alert_page(url, post_type, pd)
            if result:
                print(f"  → {post_type}: {len(result)}件")
                all_result.extend(result)
        except Exception as e:
            print(f"  → {post_type}取得失敗: {e}")

    # 重複除去（コード単位、監理>整理>改善期間の優先度）
    priority = {"監理": 0, "整理": 1, "改善期間": 2}
    by_code = {}
    for s in all_result:
        c = s["code"]
        if c not in by_code or priority.get(s["type"], 9) < priority.get(by_code[c]["type"], 9):
            by_code[c] = s
    unique = list(by_code.values())
    print(f"  → 合計 {len(unique)}件（監理{sum(1 for x in unique if x['type']=='監理')}・"
          f"整理{sum(1 for x in unique if x['type']=='整理')}・"
          f"改善期間{sum(1 for x in unique if x['type']=='改善期間')}）")
    return unique


# ── リスク判定 ───────────────────────────────────────
def calc_risk(key, total_cap, float_cap, float_ratio, listing_date):
    """
    リスクを判定。
    Returns: (risk, growth_10yr, listing_years, reasons, eff_float_cap)
    """
    c = CRITERIA[key]

    # 有効な流通株式時価総額（実値 or 総時価×0.3推定）
    if float_cap is not None:
        eff_fc = float_cap
        fc_label = f"流通時価{float_cap:.0f}億"
    elif total_cap is not None:
        eff_fc = round(total_cap * 0.3, 1)
        fc_label = f"流通時価{eff_fc:.0f}億(概算)"
    else:
        return None, False, None, [], None

    reasons = []
    risk    = None

    # ① 流通株式時価総額チェック
    if eff_fc < c["float_danger"]:
        risk = "danger"
        reasons.append(f"{fc_label} < 基準{c['float_danger']}億")
    elif eff_fc < c["float_warning"]:
        risk = "warning"
        reasons.append(f"{fc_label}（基準{c['float_danger']}億に近接）")

    # ② 流通株式比率チェック
    if float_ratio is not None and float_ratio < c["ratio_min"]:
        reasons.append(f"流通比率{float_ratio:.0f}% < 基準{c['ratio_min']}%")
        if risk is None:
            risk = "warning"

    # ③ グロース 上場10年ルール（会社全体時価総額 40億以上）
    growth_10yr  = False
    listing_years = None
    if key == "グロース" and listing_date and total_cap is not None:
        years = (TODAY - listing_date).days / 365.25
        listing_years = round(years, 1)
        if years >= 10:
            if total_cap < 40:
                growth_10yr = True
                reasons.append(f"上場{listing_years:.0f}年超・総時価{total_cap:.0f}億 < 40億")
                risk = "danger"
            elif total_cap < 80:
                growth_10yr = True
                reasons.append(f"上場{listing_years:.0f}年超・総時価{total_cap:.0f}億（40億近接）")
                if risk is None:
                    risk = "warning"

    return risk, growth_10yr, listing_years, reasons, eff_fc


# ── HTML生成 ─────────────────────────────────────────
def generate_html(stocks_data, supervision_list, generated_at):
    print("[4/4] HTMLを生成中...")
    tm = TARGET_MONTH
    pm = PREV_MONTH

    # バッジ優先度 map
    sup_map = {s["code"]: s for s in supervision_list}

    by_month = {m: {"danger": [], "warning": []} for m in range(0, 13)}
    for s in stocks_data:
        m = s["fiscal_month"] or 0
        by_month[m][s["risk"]].append(s)

    total = len(stocks_data)
    thead = ('<thead><tr><th>コード</th><th>銘柄名</th><th>市場</th>'
             '<th>流通時価総額</th><th>決算</th><th>基準日</th><th>判定・理由</th></tr></thead>')

    def stock_row(s):
        code = s["code"]
        fc   = s.get("eff_float_cap")
        cap_str = f'{fc:,.1f}億円' if fc else "取得不可"
        fi  = f'{s["fiscal_month"]}月決算' if s["fiscal_month"] else "不明"
        dl  = (f'{s["fiscal_month"]}月{calendar.monthrange(TODAY.year, s["fiscal_month"])[1]}日'
               if s["fiscal_month"] else "-")
        ms  = (s["market_str"]
               .replace("（内国株式）","")
               .replace("プライム","Prime")
               .replace("スタンダード","Std")
               .replace("グロース","Growth"))

        # 監理ポスト・改善期間バッジ
        sup_info = sup_map.get(code)
        sup_badge = ""
        if sup_info:
            type_map = {"監理": ("bs2", "監理中"), "整理": ("bd", "整理中"),
                        "改善期間": ("bkp", "改善期間中")}
            cls, lbl = type_map.get(sup_info["type"], ("bs2", sup_info["type"]))
            sup_badge = f'<span class="badge {cls}">{lbl}</span>'

        # 判定バッジ
        rb = ('<span class="badge bd">危険</span>' if s["risk"] == "danger"
              else '<span class="badge bw">注意</span>')
        if s.get("growth_10yr") and s.get("listing_years"):
            rb += f'<span class="badge bg10">上場{s["listing_years"]:.0f}年超</span>'

        # 違反理由（tooltip）
        reasons = s.get("reasons", [])
        reason_title = " / ".join(reasons) if reasons else ""
        reason_tip = (f'<div class="reason">{" | ".join(reasons)}</div>'
                      if reasons else "")

        return (f'<tr class="r{s["risk"]}" title="{reason_title}">'
                f'<td><a href="https://finance.yahoo.co.jp/quote/{code}.T" '
                f'target="_blank">{code}</a></td>'
                f'<td>{s["name"]}{sup_badge}</td><td>{ms}</td>'
                f'<td class="rc">{cap_str}</td><td>{fi}</td>'
                f'<td class="rd">{dl}</td>'
                f'<td>{rb}{reason_tip}</td></tr>')

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
                    f'<div class="tw"><table>{thead}<tbody>{rows}'
                    f'</tbody></table></div></div>')
            else:
                parts.append(
                    f'<div class="tsec"><h3 class="{cls}">{icon} {label}</h3>'
                    f'<p class="empty">該当銘柄なし</p></div>')
        return "".join(parts)

    # タブボタン生成
    tab_btns, tab_panels = [], []
    for m in range(1, 13):
        d = by_month[m]["danger"]
        w = by_month[m]["warning"]
        badges  = ""
        if d: badges += f'<span class="tbadge tbd">{len(d)}</span>'
        if w: badges += f'<span class="tbadge tbw">{len(w)}</span>'
        hl     = ('data-hl="cur"' if m == tm else
                  'data-hl="prev"' if m == pm else "")
        active = " active" if m == tm else ""
        tab_btns.append(
            f'<button class="tab{active}" onclick="showTab({m})" id="tbtn-{m}" {hl}>'
            f'{m}月{badges}</button>')
        tab_panels.append(
            f'<div class="tab-panel{active}" id="panel-{m}">{month_panel(m)}</div>')

    # 不明タブ
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

    tab_btns_html   = "\n".join(tab_btns)
    tab_panels_html = "\n".join(tab_panels)

    # 監理ポスト一覧セクション
    sup_html = ""
    if supervision_list:
        type_map = {"監理": ("bd", "監理ポスト"), "整理": ("bd", "整理ポスト"),
                    "改善期間": ("bkp", "改善期間中")}

        def sup_badge_html(t):
            cls, lbl = type_map.get(t, ("bs2", t))
            return f'<span class="badge {cls}">{lbl}</span>'

        sr = "".join(
            f'<tr><td><a href="https://finance.yahoo.co.jp/quote/{s["code"]}.T" '
            f'target="_blank">{s["code"]}</a></td>'
            f'<td>{s["name"]}</td>'
            f'<td>{sup_badge_html(s["type"])}</td>'
            f'<td>{s.get("reason","")}</td></tr>'
            for s in supervision_list)
        n_kanri = sum(1 for x in supervision_list if x["type"] == "監理")
        n_kp    = sum(1 for x in supervision_list if x["type"] == "改善期間")
        sup_html = (
            f'<section id="sup"><h2 class="csup">'
            f'🔴 監理・整理・改善期間 指定銘柄 '
            f'<span class="cnt">（監理{n_kanri}・改善期間{n_kp} / 計{len(supervision_list)}銘柄）</span></h2>'
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
.notice ul{{margin-left:1.2em;margin-top:4px}}
main{{padding:0 24px 48px;max-width:1400px;margin:0 auto}}
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
td{{padding:8px 12px;border-bottom:1px solid #f0f0f0;vertical-align:top}}
tr:last-child td{{border-bottom:none}}
.rdanger{{background:#fff5f5}}.rwarning{{background:#fffbf0}}
.rdanger:hover{{background:#ffebee}}.rwarning:hover{{background:#fff3e0}}
.rc{{text-align:right}}.rd{{font-weight:600;color:#d32f2f}}
a{{color:#1565c0;text-decoration:none}}a:hover{{text-decoration:underline}}
.badge{{display:inline-block;font-size:10px;padding:2px 6px;border-radius:3px;margin-left:4px;font-weight:600;vertical-align:middle}}
.bd{{background:#d32f2f;color:#fff}}.bw{{background:#f57c00;color:#fff}}
.bs2{{background:#880e4f;color:#fff}}.bg10{{background:#6a1b9a;color:#fff}}
.bkp{{background:#1565c0;color:#fff}}
.reason{{font-size:10px;color:#888;margin-top:3px;line-height:1.4}}
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
.tab-legend{{display:flex;gap:16px;padding:10px 24px;font-size:11px;color:#888;
             background:#fff;border-bottom:1px solid #eee;flex-wrap:wrap}}
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
<div class="notice">
⚠️ <strong>判定ロジックと免責事項</strong><br>
時価総額は <strong>流通株式時価総額</strong>（floatShares×株価）を優先使用。取得不可時は総時価×30%で概算。
<ul>
<li>✅ 検知できる基準：流通株式時価総額、流通株式比率、グロース上場10年ルール（総時価40億）</li>
<li>❌ 検知できない基準：株主数・流通株式数（東証の正式発表を参照）</li>
</ul>
投資判断の根拠にしないでください。基準日は決算月の月末最終日です。
</div>
<main>
{sup_html}
<div class="tab-wrap">
<div class="tab-label">📅 決算月で絞り込む</div>
<div class="tab-bar">
{tab_btns_html}
</div>
</div>
<div class="tab-legend">
<span><span class="tbadge tbd">N</span>危険ゾーン</span>
<span><span class="tbadge tbw">N</span>注意ゾーン</span>
<span>赤ボーダー=今月({tm}月)基準日</span>
<span>橙ボーダー=先月({pm}月)基準日</span>
<span><span class="badge bs2">監理中</span>JPX監理ポスト指定</span>
<span><span class="badge bkp">改善期間中</span>改善期間該当</span>
<span><span class="badge bg10">上場N年超</span>グロース10年ルール適用</span>
</div>
<div class="tab-panels">
{tab_panels_html}
</div>
</main>
<footer>
データ出典：<a href="https://www.jpx.co.jp/markets/statistics-equities/misc/01.html" target="_blank">東証上場銘柄一覧（JPX）</a>
・株価/浮動株：Yahoo Finance（yfinance）
・決算月：株探(kabutan.jp)補完<br>
監理/改善期間：<a href="https://www.jpx.co.jp/listing/market-alerts/supervision/index.html" target="_blank">JPX監理・整理銘柄一覧</a>
/<a href="https://www.jpx.co.jp/listing/market-alerts/improvement-period/index.html" target="_blank">改善期間該当銘柄</a><br>
本サイトは個人学習目的で作成されており、投資助言ではありません。
</footer>
<script>
function showTab(m) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
  var panel = document.getElementById('panel-' + m);
  var btn   = document.getElementById('tbtn-' + m);
  if (panel) panel.classList.add('active');
  if (btn)   btn.classList.add('active');
}}
</script>
</body></html>"""


def main():
    import pandas as pd
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    generated_at = datetime.datetime.now().strftime("%Y年%-m月%-d日 %H:%M")

    df         = fetch_stock_list()
    codes      = df["コード"].tolist()
    cap_data   = fetch_stock_data(codes)
    supervision = fetch_supervision_list()

    stocks_data = []
    for _, row_data in df.iterrows():
        code       = row_data["コード"]
        market_str = row_data["市場・商品区分"]
        info       = cap_data.get(code, {})
        total_cap  = info.get("total_cap")
        float_cap  = info.get("float_cap")
        float_ratio= info.get("float_ratio")
        fiscal_month = info.get("fiscal_month")
        listing_date = info.get("listing_date")

        key = ("プライム" if "プライム" in market_str
               else "スタンダード" if "スタンダード" in market_str
               else "グロース" if "グロース" in market_str else None)
        if not key:
            continue

        risk, growth_10yr, listing_years, reasons, eff_fc = calc_risk(
            key, total_cap, float_cap, float_ratio, listing_date)
        if not risk:
            continue

        stocks_data.append({
            "code": code, "name": row_data["銘柄名"],
            "market_str": market_str,
            "total_cap": total_cap, "float_cap": float_cap,
            "float_ratio": float_ratio, "eff_float_cap": eff_fc,
            "fiscal_month": fiscal_month, "risk": risk,
            "growth_10yr": growth_10yr, "listing_years": listing_years,
            "reasons": reasons,
        })

    stocks_data.sort(key=lambda x: (
        0 if x["fiscal_month"] == TARGET_MONTH else
        1 if x["fiscal_month"] == PREV_MONTH else 2,
        0 if x["risk"] == "danger" else 1,
        x["eff_float_cap"] or 9999
    ))

    print(f"\n対象銘柄数: {len(stocks_data)}銘柄")
    html = generate_html(stocks_data, supervision, generated_at)
    open(os.path.join(OUTPUT_DIR, "index.html"), "w", encoding="utf-8").write(html)
    json_data = {
        "generated_at": generated_at,
        "target_month": TARGET_MONTH,
        "prev_month": PREV_MONTH,
        "stocks": stocks_data,
        "supervision": supervision,
    }
    open(os.path.join(OUTPUT_DIR, "data.json"), "w", encoding="utf-8").write(
        json.dumps(json_data, ensure_ascii=False, indent=2))
    print("✅ 完了！")

if __name__ == "__main__":
    main()
