"""
東証上場維持基準 監視スクリプト
- JPX公式XLSから銘柄一覧・市場区分を取得
- Yahoo Finance Japan（非公式API）から時価総額を取得
- 上場維持基準との比較でリスク銘柄を抽出
- 今月が決算月の銘柄を優先表示
- GitHub Pages用HTMLを生成
"""

import os
import json
import datetime
import calendar
import time
import urllib.request
import urllib.parse
import subprocess
import tempfile

# ============================================================
# 定数・設定
# ============================================================

# 上場維持基準（流通株式時価総額）
# ※取得できるのは「会社全体の時価総額」なので、
#   流通株式比率を仮定して逆算（保守的に30%と仮定）
# 流通株式時価総額 ≒ 会社全体の時価総額 × 0.30
# なので: 基準の流通株式時価総額をクリアするには
#   会社全体の時価総額 ≥ 基準値 / 0.30 が目安

CRITERIA = {
    "プライム": {
        "flow_cap_billion": 100,       # 流通株式時価総額100億円以上
        "total_cap_warn_billion": 333, # 会社全体で333億円未満なら要注意（100/0.30）
        "total_cap_danger_billion": 150, # 会社全体で150億円未満なら危険ゾーン（保守的）
    },
    "スタンダード": {
        "flow_cap_billion": 10,
        "total_cap_warn_billion": 33,
        "total_cap_danger_billion": 15,
    },
    "グロース": {
        "flow_cap_billion": 5,
        "total_cap_warn_billion": 17,
        "total_cap_danger_billion": 8,
        # グロース特別: 上場10年で時価総額40億円
        "market_cap_10yr_billion": 40,
    },
}

# JPX 銘柄一覧XLSのURL（毎月第3営業日更新）
JPX_XLS_URL = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"

# 東証 監理ポスト・整理ポスト情報ページ
JPX_SUPERVISION_URL = "https://www.jpx.co.jp/rules-participants/rules/supervision/index.html"

OUTPUT_DIR = "docs"


# ============================================================
# Step 1: JPX XLSから銘柄一覧を取得
# ============================================================

def fetch_stock_list():
    """JPX公式XLSをダウンロードしてDataFrameで返す"""
    print("[1/4] JPX銘柄一覧を取得中...")
    import pandas as pd

    xls_path = "/tmp/data_j.xls"
    req = urllib.request.Request(
        JPX_XLS_URL,
        headers={"User-Agent": "Mozilla/5.0 (compatible; stock-monitor/1.0)"}
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        with open(xls_path, "wb") as f:
            f.write(resp.read())

    # LibreOffice でCSVに変換（xlrdが使えない環境向け）
    csv_path = "/tmp/data_j.csv"
    subprocess.run(
        ["python3", "/usr/local/lib/python3.12/dist-packages/scripts/office/soffice.py",
         "--headless", "--convert-to", "csv", xls_path, "--outdir", "/tmp/"],
        check=False, capture_output=True
    )

    # sofficeラッパーがない場合は直接呼ぶ
    if not os.path.exists(csv_path):
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "csv", xls_path, "--outdir", "/tmp/"],
            check=True, capture_output=True, timeout=60
        )

    df = pd.read_csv(csv_path, encoding="utf-8")
    # 内国普通株式のみ（ETF・REIT等を除外）
    df = df[df["市場・商品区分"].str.contains("内国株式", na=False)].copy()
    df["コード"] = df["コード"].astype(str).str.strip()
    print(f"  → {len(df)}銘柄を取得")
    return df


# ============================================================
# Step 2: Yahoo Finance Japan から時価総額・決算月を取得
# ============================================================

def fetch_market_cap_yahoo(codes):
    """
    Yahoo Finance Japan の非公式APIから時価総額・決算月を取得。
    100件ずつバッチ処理。
    返り値: {コード: {"market_cap": 億円, "fiscal_month": 決算月(1-12 or None)}}
    """
    print("[2/4] 時価総額・決算月を取得中（Yahoo Finance Japan）...")
    results = {}
    batch_size = 50
    codes_list = list(codes)

    for i in range(0, len(codes_list), batch_size):
        batch = codes_list[i:i+batch_size]
        # Yahoo Finance のサマリーAPIを使用
        # 例: https://query1.finance.yahoo.com/v8/finance/chart/1301.T
        for code in batch:
            ticker = f"{code}.T"
            url = (
                f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{ticker}"
                f"?modules=summaryDetail%2CdefaultKeyStatistics%2CassetProfile"
            )
            try:
                req = urllib.request.Request(
                    url,
                    headers={
                        "User-Agent": "Mozilla/5.0",
                        "Accept": "application/json",
                    }
                )
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = json.loads(resp.read().decode("utf-8"))

                summary = data.get("quoteSummary", {}).get("result", [{}])[0]

                # 時価総額（円 → 億円）
                market_cap_raw = (
                    summary.get("summaryDetail", {})
                    .get("marketCap", {})
                    .get("raw", None)
                )
                market_cap_billion = round(market_cap_raw / 1e8, 1) if market_cap_raw else None

                # 決算月（fiscalYearEnd例: "March" → 3）
                fiscal_year_end = (
                    summary.get("summaryDetail", {})
                    .get("fiscalYearEnd", {})
                    .get("fmt", None)
                )
                fiscal_month = parse_fiscal_month(fiscal_year_end)

                results[code] = {
                    "market_cap": market_cap_billion,
                    "fiscal_month": fiscal_month,
                }

            except Exception as e:
                results[code] = {"market_cap": None, "fiscal_month": None}

        # レートリミット対策
        time.sleep(0.5)
        print(f"  → {min(i+batch_size, len(codes_list))}/{len(codes_list)}件処理済み")

    return results


def parse_fiscal_month(fmt):
    """Yahoo Financeの決算月文字列（例: "March"）を数値に変換"""
    if not fmt:
        return None
    month_map = {
        "january": 1, "february": 2, "march": 3, "april": 4,
        "may": 5, "june": 6, "july": 7, "august": 8,
        "september": 9, "october": 10, "november": 11, "december": 12,
    }
    return month_map.get(fmt.lower(), None)


# ============================================================
# Step 3: 東証 監理ポスト・整理ポスト情報を取得
# ============================================================

def fetch_supervision_list():
    """
    東証の監理ポスト・整理ポスト指定銘柄をスクレイピング。
    返り値: [{"code": "1234", "name": "〇〇株式会社", "type": "監理", "reason": "..."}]
    """
    print("[3/4] 監理ポスト・整理ポスト情報を取得中...")
    try:
        from html.parser import HTMLParser

        req = urllib.request.Request(
            JPX_SUPERVISION_URL,
            headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="replace")

        # 簡易パース：表からコード・銘柄名・理由を抽出
        supervision = []
        # テーブル行を正規表現で抽出
        import re
        # <td>内のコード（4桁数字）を探す
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL)
        for row in rows:
            cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
            cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
            if len(cells) >= 2:
                # コードっぽいものを探す
                code_candidate = cells[0].replace("\u3000", "").strip()
                if re.match(r'^\d{4}$', code_candidate):
                    supervision.append({
                        "code": code_candidate,
                        "name": cells[1] if len(cells) > 1 else "",
                        "type": "監理" if "監理" in html[html.find(code_candidate)-500:html.find(code_candidate)] else "整理",
                        "reason": cells[2] if len(cells) > 2 else "",
                    })

        print(f"  → 監理・整理ポスト {len(supervision)}件を取得")
        return supervision

    except Exception as e:
        print(f"  → 監視ポスト取得失敗（{e}）。空リストで続行。")
        return []


# ============================================================
# Step 4: 判定ロジック
# ============================================================

def classify_risk(row, market_cap_billion):
    """
    銘柄のリスクレベルを判定。
    返り値: "danger" | "warning" | "ok" | "unknown"
    """
    market = row["市場・商品区分"]
    if market_cap_billion is None:
        return "unknown"

    if "プライム" in market:
        key = "プライム"
    elif "スタンダード" in market:
        key = "スタンダード"
    elif "グロース" in market:
        key = "グロース"
    else:
        return "unknown"

    c = CRITERIA[key]
    if market_cap_billion < c["total_cap_danger_billion"]:
        return "danger"
    elif market_cap_billion < c["total_cap_warn_billion"]:
        return "warning"
    else:
        return "ok"


def is_this_month_fiscal(fiscal_month):
    """今月が決算月かどうか"""
    if fiscal_month is None:
        return False
    return fiscal_month == datetime.date.today().month


def get_deadline_date(fiscal_month):
    """決算月の月末最終日を返す"""
    if fiscal_month is None:
        return None
    today = datetime.date.today()
    year = today.year
    # 今年の該当月の月末
    last_day = calendar.monthrange(year, fiscal_month)[1]
    return datetime.date(year, fiscal_month, last_day)


# ============================================================
# Step 5: HTML生成
# ============================================================

def generate_html(stocks_data, supervision_list, generated_at):
    """GitHub Pages用HTMLを生成"""
    print("[4/4] HTMLを生成中...")

    today = datetime.date.today()
    this_month = today.month
    this_month_ja = f"{this_month}月"

    # リスク別に分類
    danger_this_month = []
    warning_this_month = []
    danger_other = []
    warning_other = []

    for s in stocks_data:
        risk = s["risk"]
        this_month_fiscal = s["this_month_fiscal"]
        if risk == "danger":
            if this_month_fiscal:
                danger_this_month.append(s)
            else:
                danger_other.append(s)
        elif risk == "warning":
            if this_month_fiscal:
                warning_this_month.append(s)
            else:
                warning_other.append(s)

    # 監理・整理ポストのコードセット
    supervision_codes = {s["code"] for s in supervision_list}

    def row_html(s, highlight_class=""):
        code = s["code"]
        name = s["name"]
        market = s["market_short"]
        cap = f'{s["market_cap"]:,.0f}億円' if s["market_cap"] else "取得不可"
        fiscal = f'{s["fiscal_month"]}月決算' if s["fiscal_month"] else "不明"
        deadline = s["deadline"].strftime("%-m月%-d日") if s["deadline"] else "-"
        in_supervision = code in supervision_codes
        supervision_badge = '<span class="badge badge-supervision">監理中</span>' if in_supervision else ""
        risk_label = {
            "danger": '<span class="badge badge-danger">危険</span>',
            "warning": '<span class="badge badge-warning">注意</span>',
        }.get(s["risk"], "")

        yahoo_url = f"https://finance.yahoo.co.jp/quote/{code}.T"
        return f"""
        <tr class="{highlight_class}">
          <td><a href="{yahoo_url}" target="_blank" rel="noopener">{code}</a></td>
          <td>{name}{supervision_badge}</td>
          <td>{market}</td>
          <td class="cap-cell">{cap}</td>
          <td>{fiscal}</td>
          <td class="deadline-cell">{deadline}</td>
          <td>{risk_label}</td>
        </tr>"""

    def table_html(title, items, section_id, color_class):
        if not items:
            return f'<section id="{section_id}"><h2 class="{color_class}">{title}</h2><p class="empty">該当銘柄なし</p></section>'
        rows = "".join(row_html(s, color_class) for s in items)
        count = len(items)
        return f"""
    <section id="{section_id}">
      <h2 class="{color_class}">{title} <span class="count">({count}銘柄)</span></h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>コード</th><th>銘柄名</th><th>市場</th>
              <th>時価総額</th><th>決算</th><th>基準日</th><th>判定</th>
            </tr>
          </thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
    </section>"""

    # 監理・整理ポストテーブル
    if supervision_list:
        sup_rows = "".join(f"""
        <tr>
          <td>{s['code']}</td>
          <td>{s['name']}</td>
          <td><span class="badge badge-{'danger' if s['type']=='整理' else 'supervision'}">{s['type']}ポスト</span></td>
          <td>{s.get('reason','')}</td>
        </tr>""" for s in supervision_list)
        supervision_section = f"""
    <section id="supervision">
      <h2 class="color-supervision">🔴 現在の監理・整理ポスト指定銘柄 <span class="count">({len(supervision_list)}銘柄)</span></h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>コード</th><th>銘柄名</th><th>種別</th><th>指定理由</th></tr></thead>
          <tbody>{sup_rows}</tbody>
        </table>
      </div>
    </section>"""
    else:
        supervision_section = ""

    total_alert = len(danger_this_month) + len(warning_this_month) + len(danger_other) + len(warning_other)

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>東証 上場維持基準 監視サイト</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif;
      background: #f4f5f7;
      color: #1a1a2e;
      font-size: 14px;
      line-height: 1.6;
    }}
    header {{
      background: #1a1a2e;
      color: #fff;
      padding: 20px 24px;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
      flex-wrap: wrap;
      gap: 8px;
    }}
    header h1 {{ font-size: 20px; font-weight: 700; letter-spacing: 0.02em; }}
    header .meta {{ font-size: 12px; opacity: 0.7; }}
    .summary-bar {{
      background: #fff;
      border-bottom: 1px solid #e0e0e0;
      padding: 12px 24px;
      display: flex;
      gap: 24px;
      flex-wrap: wrap;
      align-items: center;
    }}
    .summary-bar .label {{ font-size: 12px; color: #666; }}
    .summary-bar .num {{ font-size: 22px; font-weight: 700; margin-left: 4px; }}
    .num.danger {{ color: #d32f2f; }}
    .num.warning {{ color: #f57c00; }}
    .notice {{
      background: #fff8e1;
      border-left: 4px solid #f9a825;
      margin: 16px 24px;
      padding: 12px 16px;
      font-size: 12px;
      color: #555;
      border-radius: 0 4px 4px 0;
    }}
    main {{ padding: 0 24px 48px; max-width: 1200px; margin: 0 auto; }}
    section {{ margin-top: 28px; }}
    h2 {{
      font-size: 16px;
      font-weight: 700;
      padding: 10px 14px;
      border-radius: 6px 6px 0 0;
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    h2.color-danger {{ background: #ffebee; color: #b71c1c; border: 1px solid #ef9a9a; }}
    h2.color-warning {{ background: #fff3e0; color: #bf360c; border: 1px solid #ffcc80; }}
    h2.color-supervision {{ background: #fce4ec; color: #880e4f; border: 1px solid #f48fb1; }}
    .count {{ font-size: 13px; font-weight: 400; opacity: 0.8; }}
    .table-wrap {{ overflow-x: auto; }}
    table {{
      width: 100%;
      border-collapse: collapse;
      background: #fff;
      border: 1px solid #e0e0e0;
      border-top: none;
    }}
    th {{
      background: #f5f5f5;
      padding: 8px 12px;
      text-align: left;
      font-size: 12px;
      color: #555;
      border-bottom: 1px solid #ddd;
      white-space: nowrap;
    }}
    td {{
      padding: 8px 12px;
      border-bottom: 1px solid #f0f0f0;
      vertical-align: middle;
    }}
    tr:last-child td {{ border-bottom: none; }}
    tr.color-danger {{ background: #fff5f5; }}
    tr.color-warning {{ background: #fffbf0; }}
    tr.color-danger:hover {{ background: #ffebee; }}
    tr.color-warning:hover {{ background: #fff3e0; }}
    .cap-cell {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .deadline-cell {{ font-weight: 600; color: #d32f2f; }}
    a {{ color: #1565c0; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .badge {{
      display: inline-block;
      font-size: 10px;
      padding: 2px 6px;
      border-radius: 3px;
      margin-left: 4px;
      font-weight: 600;
      vertical-align: middle;
    }}
    .badge-danger {{ background: #d32f2f; color: #fff; }}
    .badge-warning {{ background: #f57c00; color: #fff; }}
    .badge-supervision {{ background: #880e4f; color: #fff; }}
    .empty {{ padding: 16px; background: #fff; border: 1px solid #e0e0e0; border-top: none; color: #999; font-size: 13px; }}
    footer {{
      text-align: center;
      padding: 24px;
      font-size: 11px;
      color: #999;
      border-top: 1px solid #e0e0e0;
      margin-top: 40px;
    }}
    @media (max-width: 600px) {{
      header, .summary-bar, main {{ padding-left: 12px; padding-right: 12px; }}
      .notice {{ margin-left: 12px; margin-right: 12px; }}
      table {{ font-size: 12px; }}
      td, th {{ padding: 6px 8px; }}
    }}
  </style>
</head>
<body>
<header>
  <div>
    <h1>📊 東証 上場維持基準 監視サイト</h1>
    <div>2025年3月〜厳格化された上場維持基準に注意が必要な銘柄を自動抽出</div>
  </div>
  <div class="meta">最終更新: {generated_at}<br>毎日自動更新（GitHub Actions）</div>
</header>

<div class="summary-bar">
  <div>
    <span class="label">今月基準日・危険ゾーン</span>
    <span class="num danger">{len(danger_this_month)}</span>
    <span class="label">銘柄</span>
  </div>
  <div>
    <span class="label">今月基準日・注意ゾーン</span>
    <span class="num warning">{len(warning_this_month)}</span>
    <span class="label">銘柄</span>
  </div>
  <div>
    <span class="label">監視対象合計（全月）</span>
    <span class="num" style="color:#1565c0">{total_alert}</span>
    <span class="label">銘柄</span>
  </div>
</div>

<div class="notice">
  ⚠️ <strong>免責事項：</strong>
  表示される時価総額は「会社全体の時価総額」（株価 × 発行済株式数）の推定値です。
  上場維持基準の判定に使われる「流通株式時価総額」とは異なります。
  流通株式比率を30%と仮定したおおよその目安として参考にしてください。
  投資判断の根拠にはしないでください。
  基準日は各銘柄の決算月の月末最終日です。
</div>

<main>
  {table_html(f"🔴 今月({this_month_ja})が基準日・危険ゾーン（時価総額が基準の半分以下の目安）", danger_this_month, "danger-this-month", "color-danger")}
  {table_html(f"🟠 今月({this_month_ja})が基準日・注意ゾーン（時価総額が基準付近の目安）", warning_this_month, "warning-this-month", "color-warning")}
  {supervision_section}
  {table_html("⚠️ 危険ゾーン（来月以降が基準日）", danger_other, "danger-other", "color-danger")}
  {table_html("💛 注意ゾーン（来月以降が基準日）", warning_other, "warning-other", "color-warning")}
</main>

<footer>
  データ出典：<a href="https://www.jpx.co.jp/markets/statistics-equities/misc/01.html" target="_blank">東証上場銘柄一覧（JPX）</a>、
  時価総額：Yahoo Finance Japan（非公式）<br>
  本サイトは個人学習目的で作成されており、投資助言ではありません。
</footer>
</body>
</html>"""

    return html


# ============================================================
# メイン処理
# ============================================================

def main():
    import pandas as pd

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    generated_at = datetime.datetime.now().strftime("%Y年%-m月%-d日 %H:%M")

    # Step 1: 銘柄一覧
    df = fetch_stock_list()

    # Step 2: 時価総額・決算月
    codes = df["コード"].tolist()
    cap_data = fetch_market_cap_yahoo(codes)

    # Step 3: 監理ポスト
    supervision = fetch_supervision_list()

    # Step 4: 統合・判定
    stocks_data = []
    for _, row in df.iterrows():
        code = row["コード"]
        market_str = row["市場・商品区分"]
        cap_info = cap_data.get(code, {})
        market_cap = cap_info.get("market_cap")
        fiscal_month = cap_info.get("fiscal_month")

        risk = classify_risk(row, market_cap)
        if risk not in ("danger", "warning"):
            continue  # okとunknownは除外

        if "プライム" in market_str:
            market_short = "Prime"
        elif "スタンダード" in market_str:
            market_short = "Standard"
        elif "グロース" in market_str:
            market_short = "Growth"
        else:
            market_short = market_str

        deadline = get_deadline_date(fiscal_month)

        stocks_data.append({
            "code": code,
            "name": row["銘柄名"],
            "market_short": market_short,
            "market_cap": market_cap,
            "fiscal_month": fiscal_month,
            "this_month_fiscal": is_this_month_fiscal(fiscal_month),
            "deadline": deadline,
            "risk": risk,
        })

    # 今月・危険順でソート
    stocks_data.sort(key=lambda x: (
        not x["this_month_fiscal"],
        x["risk"] != "danger",
        x["market_cap"] or 9999
    ))

    # Step 5: HTML生成
    html = generate_html(stocks_data, supervision, generated_at)

    output_path = os.path.join(OUTPUT_DIR, "index.html")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    # JSONも出力（将来の拡張用）
    json_data = {
        "generated_at": generated_at,
        "stocks": stocks_data,
        "supervision": supervision,
    }
    # dateオブジェクトをstrに変換
    for s in json_data["stocks"]:
        if s["deadline"]:
            s["deadline"] = s["deadline"].isoformat()

    with open(os.path.join(OUTPUT_DIR, "data.json"), "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 完了！")
    print(f"  対象銘柄数: {len(stocks_data)}銘柄")
    print(f"  出力先: {output_path}")


if __name__ == "__main__":
    main()
