"""
新規プロジェクト作成スクリプト
使い方: python 09_system/new_project.py
"""

import os
import sys
from datetime import date

# Windows でのエンコーディング設定
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")

# パス設定
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLIENTS_DIR = os.path.join(BASE_DIR, "03_clients")
TEMPLATE_PATH = os.path.join(BASE_DIR, "09_system", "config", "PROJECT_TEMPLATE.md")


def list_dirs(path):
    """フォルダ一覧を返す（隠しフォルダ除く）"""
    return sorted([
        d for d in os.listdir(path)
        if os.path.isdir(os.path.join(path, d)) and not d.startswith(".")
    ])


def select_from_list(label: str, options: list[str], allow_custom: bool = True) -> str:
    """一覧から選択させる"""
    print(f"\n{label}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    if allow_custom:
        print(f"  {len(options)+1}. 直接入力する")

    while True:
        choice = input("番号を入力: ").strip()
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(options):
                return options[idx - 1]
            if allow_custom and idx == len(options) + 1:
                val = input("入力してください: ").strip()
                return val
        print("もう一度入力してください")
    return ""  # 到達しないが型チェック用


def ask(prompt, default=""):
    """デフォルト値付き入力"""
    suffix = f" [{default}]" if default else ""
    val = input(f"{prompt}{suffix}: ").strip()
    return val if val else default


def create_project():
    print("=" * 40)
    print("  新規プロジェクト作成")
    print("=" * 40)

    # クライアント選択
    clients = list_dirs(CLIENTS_DIR)
    if not clients:
        print("エラー: 03_clients/ にクライアントフォルダがありません")
        return
    client_name = select_from_list("クライアントを選択してください:", clients)
    client_dir = os.path.join(CLIENTS_DIR, client_name)

    # サービス区分選択（projects/ 以外のフォルダがあれば）
    sub_dirs = [d for d in list_dirs(client_dir) if d != "projects"]
    if sub_dirs:
        service = select_from_list("サービス区分を選択してください:", sub_dirs)
        project_parent = os.path.join(client_dir, service)
    else:
        service = ""
        project_parent = os.path.join(client_dir, "projects")

    # プロジェクト名
    project_name = input("\nプロジェクト名を入力してください: ").strip()
    if not project_name:
        print("エラー: プロジェクト名は必須です")
        return

    # 各種情報入力
    today = date.today().strftime("%Y-%m-%d")
    client_contact = ask("クライアント担当者名")
    internal_contact = ask("社内担当者名", "npino")
    start_date = ask("開始日", today)
    end_date = ask("終了予定日", "未定")

    # ステータス
    status_options = ["検討中", "進行中", "完了", "停止"]
    status = select_from_list("ステータスを選択してください:", status_options, allow_custom=False)

    # フォルダ作成
    project_dir = os.path.join(project_parent, project_name)
    if os.path.exists(project_dir):
        print(f"\nエラー: 既に同名のフォルダが存在します → {project_dir}")
        return
    os.makedirs(project_dir, exist_ok=True)

    # テンプレート読み込み
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        content = f.read()

    # 基本情報を埋め込む
    replacements = {
        "| プロジェクト名 | |": f"| プロジェクト名 | {project_name} |",
        "| クライアント名 | |": f"| クライアント名 | {client_name} |",
        "| サービス区分 | （例：プロプレミアムTEAM / マンツーマンチーム） |":
            f"| サービス区分 | {service} |",
        "| 開始日 | |": f"| 開始日 | {start_date} |",
        "| 終了予定日 | |": f"| 終了予定日 | {end_date} |",
        "| ステータス | 検討中 / 進行中 / 完了 / 停止 |": f"| ステータス | {status} |",
        "| 社内担当者 | |": f"| 社内担当者 | {internal_contact} |",
        "| クライアント担当者 | |": f"| クライアント担当者 | {client_contact} |",
        "### YYYY-MM-DD": f"### {today}",
    }
    for old, new in replacements.items():
        content = content.replace(old, new)

    # ファイル書き出し
    output_path = os.path.join(project_dir, "PROJECT.md")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    rel_path = os.path.relpath(project_dir, BASE_DIR).replace("\\", "/")
    print(f"\n完了！プロジェクトを作成しました")
    print(f"  場所: {rel_path}")
    print(f"  ファイル: PROJECT.md")


if __name__ == "__main__":
    create_project()
