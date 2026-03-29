"""
会員ロースター・KPI進捗をGoogle Sheetsカルテから同期するスクリプト

使い方:
  python sync_member_data.py          # member_roster.md + kpi_progress.md を最新化
"""

import sys
import io
import re
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# Windows環境のUTF-8対応
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# --- 設定 ---
BASE_DIR = Path(__file__).parent.parent.parent
CONFIG_DIR = BASE_DIR / "09_system" / "config"
SA_FILE = CONFIG_DIR / "google_service_account.json"
SPREADSHEET_ID = "1EHKgmE1d7T5N9GbT_QJV72Q8Dh91iNysHHkz6LEWJy4"
ROSTER_PATH = BASE_DIR / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "会員管理" / "member_roster.md"
KPI_PATH = BASE_DIR / "03_clients" / "SIFTAI" / "プロプレミアムTEAM" / "コーチング" / "kpi_progress.md"


def get_token():
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import Request
    creds = Credentials.from_service_account_file(
        str(SA_FILE), scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    creds.refresh(Request())
    return creds.token


def fetch(token, sheet, rng="A:Z"):
    import requests as req
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/{sheet}!{rng}"
    r = req.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    rows = r.json().get("values", [])
    return rows[0] if rows else [], rows[1:] if len(rows) > 1 else []


def safe(r, i):
    return r[i].strip() if len(r) > i and r[i] else ""


def status_label(diary, fb, start_date_str):
    if start_date_str:
        for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                sd = datetime.strptime(start_date_str.strip(), fmt)
                if sd > datetime.now():
                    return "未開始"
                break
            except ValueError:
                continue
    if diary <= 2:
        return "要緊急フォロー"
    if diary <= 9 or fb == 0:
        return "要注意"
    return "アクティブ"


def safe_int(r, i):
    """数値列を安全にintに変換"""
    val = safe(r, i)
    if not val:
        return 0
    nums = re.findall(r"\d+", val)
    return int(nums[0]) if nums else 0


def main():
    print("Google Sheets接続中...")
    token = get_token()

    # PRO列: [0]ID [1]名前 [2]カルテURL [3]歩き方URL [4]開始日 [5]終了日 [6]期数 [7]日報 [8]週報 [9]FB会
    _, pro_rows = fetch(token, "DB_Pro-Members", "A:J")
    # PREMIUM列: [0]ID [1]名前 [2]カルテURL [3]シートURL [4]歩き方URL [5]開始日 [6]終了日 [7]日報 [8]週報 [9]FB会
    _, prem_rows = fetch(token, "DB_Premium-Members", "A:J")
    # 週報列: [0]タイムスタンプ [1]氏名 [2]日付 [3]稼働時間 [4]成果物 [5]フェーズ [6]うまくいった点 [7]課題 [8]来週目標 [9]マ��タイズ成果
    _, weekly_rows = fetch(token, "DB_Weekly-Report", "A:J")

    print(f"PRO: {len(pro_rows)}名, PREMIUM: {len(prem_rows)}名, 週報: {len(weekly_rows)}件")

    today = datetime.now().strftime("%Y-%m-%d")

    pro_names = set(safe(r, 1) for r in pro_rows if safe(r, 1))
    prem_names = set(safe(r, 1) for r in prem_rows if safe(r, 1))

    # ==================== member_roster.md ====================
    # PRO: Sheetsの累計値をそのまま使用
    pro_lines = []
    pro_active = pro_caution = pro_urgent = pro_not_started = 0
    for i, r in enumerate(pro_rows):
        mid = safe(r, 0)
        name = safe(r, 1)
        if not name:
            continue
        kisu = safe(r, 6)   # 期数
        start = safe(r, 4)  # 開始日
        end = safe(r, 5)    # 終了日
        d = safe_int(r, 7)  # 日報（Sheets累計）
        w = safe_int(r, 8)  # 週報���Sheets累計）
        f = safe_int(r, 9)  # FB会（Sheets累計）
        is_pre = "○" if name in prem_names else "—"
        st = status_label(d, f, start)
        if st == "アクティブ":
            pro_active += 1
        elif st == "要注意":
            pro_caution += 1
        elif st == "要緊急フォロー":
            pro_urgent += 1
        elif st == "��開始":
            pro_not_started += 1
        pro_lines.append(
            f"| {i+1} | {mid} | {name} | {kisu} | {start} | {end} | {d} | {w} | {f} | {is_pre} | {st} |"
        )

    # PREMIUM: Sheetsの累計値をそのまま使用
    prem_lines = []
    prem_active = prem_caution = prem_urgent = 0
    for i, r in enumerate(prem_rows):
        mid = safe(r, 0)
        name = safe(r, 1)
        if not name:
            continue
        start = safe(r, 5)  # 開始日
        end = safe(r, 6)    # 終了日
        d = safe_int(r, 7)  # 日報（Sheets累計）
        is_pro = "○" if name in pro_names else "—"
        st = status_label(d, 1, start)  # PREMIUMはFB会必須でない
        if st == "アクティブ":
            prem_active += 1
        elif st == "要注意":
            prem_caution += 1
        elif st == "要緊急フォロー":
            prem_urgent += 1
        prem_lines.append(
            f"| {i+1} | {mid} | {name} | {start} | {end} | {d} | {is_pro} | {st} |"
        )

    both = sorted(pro_names & prem_names)

    roster_md = f"""# 会員ロースター

最終更新：{today}（Google Sheets カルテより自動同期）

データソース: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/

---

## PRO会員（{len(pro_lines)}名）

| # | ID | 名前 | 期数 | 開始日 | 終了日 | 日報 | 週報 | FB会 | PRE併用 | 状態 |
|---|---|---|---|---|---|---:|---:|---:|:---:|---|
""" + "\n".join(pro_lines) + f"""

---

## PREMIUM会員（{len(prem_lines)}名）

| # | ID | 名前 | 開始日 | 終了日 | 日報 | PRO併用 | 状態 |
|---|---|---|---|---|---:|:---:|---|
""" + "\n".join(prem_lines) + """

---

## PRO/PREMIUM 併用メンバー

""" + "\n".join(f"- {n}" for n in both) + f"""

---

## サマリー

| 指標 | PRO | PREMIUM |
|---|---|---|
| 総在籍数 | {len(pro_lines)}名 | {len(prem_lines)}名 |
| アクティブ | {pro_active}名 | {prem_active}名 |
| 要注意 | {pro_caution}名 | {prem_caution}名 |
| 要緊急フォロー | {pro_urgent}名 | {prem_urgent}名 |
| 未開始 | {pro_not_started}名 | — |

---

## 判定基準

- **アクティブ**: 累計日報10回以上 かつ FB会参加あり
- **要注意**: 累計日報3〜9回 または FB会ゼロ
- **要緊急フォロー**: 累計日報2回以下
- **未開始**: 開始日が未到来
"""

    ROSTER_PATH.write_text(roster_md, encoding="utf-8")
    print(f"✅ member_roster.md 更新完了（{today}）")

    # ==================== kpi_progress.md ====================
    # 週報列: [5]フェーズ [9]マネタイズ成果
    monetize = {}
    for r in weekly_rows:
        if len(r) > 1:
            name = r[1]
            amount_str = safe(r, 9)  # マネタイズ成果
            content = safe(r, 4)     # 成果物・アウトプット
            phase = safe(r, 5)       # フェーズ
            if amount_str:
                nums = re.findall(r"\d+", amount_str.replace(",", ""))
                if nums:
                    try:
                        amt = int(nums[0])
                        if name not in monetize:
                            monetize[name] = {"amount": 0, "content": "", "phase": phase}
                        monetize[name]["amount"] += amt
                        if content:
                            monetize[name]["content"] = content
                        if phase:
                            monetize[name]["phase"] = phase
                    except ValueError:
                        pass

    def get_plan(name):
        in_pro = name in pro_names
        in_pre = name in prem_names
        if in_pro and in_pre:
            return "PRO+PRE"
        elif in_pro:
            return "PRO"
        elif in_pre:
            return "PREMIUM"
        return "—"

    ranked = sorted(monetize.items(), key=lambda x: x[1]["amount"], reverse=True)
    ranked = [(n, d) for n, d in ranked if d["amount"] > 0]
    total_amount = sum(d["amount"] for _, d in ranked)

    kpi_lines = []
    for i, (name, d) in enumerate(ranked):
        kpi_lines.append(
            f'| {i+1} | {name} | {get_plan(name)} | {d["phase"]} | {d["amount"]:,}円 | {d["content"]} |'
        )

    urgent_pro_lines = []
    for r in pro_rows:
        name = safe(r, 1)
        if not name:
            continue
        d = safe_int(r, 7)  # 日報（Sheets累計）
        f = safe_int(r, 9)  # FB会（Sheets累計）
        if d <= 2:
            urgent_pro_lines.append(f"| {name} | {d} | {f} | 要緊急フォロー |")

    urgent_prem_lines = []
    for r in prem_rows:
        name = safe(r, 1)
        if not name:
            continue
        d = safe_int(r, 7)  # 日報（Sheets累計）
        if d <= 2:
            urgent_prem_lines.append(f"| {name} | {d} | 要緊急フォロー |")

    top5 = "\n".join(
        f'| {n} | {get_plan(n)} | {d["amount"]:,}円 — {d["content"]} |' for n, d in ranked[:5]
    )

    kpi_md = f"""# 会員KPI進捗ダッシュボード

最終更新：{today}（Google Sheets カルテより自動同期）

データソース: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/

---

## マネタイズ成果ランキング

| 順位 | 名前 | プラン | フェーズ | 累計額 | 内容 |
|:---:|---|---|---|---:|---|
""" + "\n".join(kpi_lines) + f"""

合計: **{total_amount:,}円**（{len(ranked)}名）

---

## 要緊急フォロー対象

### PRO

| 名前 | 累計日報 | 累計FB | 状態 |
|---|---:|---:|---|
""" + "\n".join(urgent_pro_lines) + """

### PREMIUM

| 名前 | 累計日報 | 状態 |
|---|---:|---|
""" + "\n".join(urgent_prem_lines) + f"""

---

## 成果事例（表彰・共有候補）

| 名前 | プラン | 成果内容 |
|---|---|---|
""" + top5 + """
"""

    KPI_PATH.write_text(kpi_md, encoding="utf-8")
    print(f"✅ kpi_progress.md 更新完了（{today}）")
    print("全ファイル同期完了")


if __name__ == "__main__":
    main()
