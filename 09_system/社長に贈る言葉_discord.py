"""
毎朝Discordに名著からの一言を配信するスクリプト
- 日替わりで名言を選択（日付ベースでローテーション）
- Discord Webhookで送信
"""

import os
import sys
import json
import random
import requests
from datetime import datetime, timezone, timedelta
from pathlib import Path

# Windows環境でのUTF-8出力対応
os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# === 設定 ===
WEBHOOK_URL = "https://discord.com/api/webhooks/1480602170368856275/DCDEaDwYhIIH0xGioPjanNbTocElv0FxtbnGo7HTqSDfpR_bn_IPO6Rm2PNhC7p0gUB4"
SCRIPT_DIR = Path(__file__).resolve().parent
HISTORY_PATH = SCRIPT_DIR / "config" / "quote_history.json"

JST = timezone(timedelta(hours=9))
TODAY = datetime.now(JST)

# === 名言データベース ===
# カテゴリ: action（行動）, mindset（心構え）, growth（成長）, leadership（リーダーシップ）, creativity（創造性）

QUOTES = [
    # ── 行動・実行 ──
    {"quote": "完璧を目指すよりまず終わらせろ。", "author": "マーク・ザッカーバーグ", "book": "（Facebook社内標語）", "category": "action"},
    {"quote": "1%の改善を毎日続ければ、1年後には37倍になる。", "author": "ジェームズ・クリアー", "book": "『Atomic Habits（複利で伸びる1つの習慣）』", "category": "action"},
    {"quote": "始める前にすべてを知る必要はない。始めることで学ぶのだ。", "author": "リード・ホフマン", "book": "『ALLIANCE』", "category": "action"},
    {"quote": "今日やれることを明日に延ばすな。", "author": "ベンジャミン・フランクリン", "book": "『フランクリン自伝』", "category": "action"},
    {"quote": "小さく始めて、素早く失敗し、大きく学べ。", "author": "エリック・リース", "book": "『リーン・スタートアップ』", "category": "action"},
    {"quote": "計画は役に立たないが、計画することは不可欠だ。", "author": "ドワイト・D・アイゼンハワー", "book": "", "category": "action"},
    {"quote": "世界を変えたいなら、まずベッドメイキングから始めよ。", "author": "ウィリアム・マクレイヴン", "book": "『Make Your Bed』", "category": "action"},
    {"quote": "やる気がなくても行動せよ。行動がやる気を生む。", "author": "マーク・マンソン", "book": "『「気がつきすぎて疲れる」が驚くほどなくなる』", "category": "action"},
    {"quote": "アウトプットしないインプットは、ただの自己満足だ。", "author": "樺沢紫苑", "book": "『アウトプット大全』", "category": "action"},
    {"quote": "プロとアマの違いは、やる気がない日でもやるかどうかだ。", "author": "スティーブン・プレスフィールド", "book": "『やりとげる力』", "category": "action"},
    {"quote": "20%の努力が80%の成果を生む。大事なのは、どの20%かを見極めることだ。", "author": "リチャード・コッチ", "book": "『80/20の法則』", "category": "action"},
    {"quote": "最も重要なことを、最も重要でないことの犠牲にしてはならない。", "author": "ゲーテ", "book": "", "category": "action"},
    {"quote": "朝の1時間は、夜の3時間に匹敵する。", "author": "ロビン・シャルマ", "book": "『5時起きの習慣』", "category": "action"},

    # ── 心構え・マインドセット ──
    {"quote": "人生で大切なのは、自分が何を持っているかではなく、持っているもので何をするかだ。", "author": "アドラー（岸見一郎 訳）", "book": "『嫌われる勇気』", "category": "mindset"},
    {"quote": "あなたの時間は限られている。他人の人生を生きて無駄にしてはいけない。", "author": "スティーブ・ジョブズ", "book": "（スタンフォード大学スピーチ）", "category": "mindset"},
    {"quote": "困難の中にこそ、チャンスがある。", "author": "アルベルト・アインシュタイン", "book": "", "category": "mindset"},
    {"quote": "成功とは、情熱を失わずに失敗から失敗へと進む能力のことだ。", "author": "ウィンストン・チャーチル", "book": "", "category": "mindset"},
    {"quote": "人は変われる。そして、人はいつでも変われる。", "author": "アドラー（岸見一郎 訳）", "book": "『幸せになる勇気』", "category": "mindset"},
    {"quote": "今この瞬間を生きることだけが、本当の人生だ。", "author": "エックハルト・トール", "book": "『さとりをひらくと人生はシンプルで楽になる』", "category": "mindset"},
    {"quote": "自分の機嫌は自分でとれ。他人に期待するな。", "author": "みうらじゅん", "book": "『「ない仕事」の作り方』", "category": "mindset"},
    {"quote": "不安は未来に対する想像力の誤用である。", "author": "デール・カーネギー", "book": "『道は開ける』", "category": "mindset"},
    {"quote": "幸せだから感謝するのではない。感謝するから幸せなのだ。", "author": "デヴィッド・スタインドル＝ラスト", "book": "", "category": "mindset"},
    {"quote": "他人と比較するな。昨日の自分と比較しろ。", "author": "ジョーダン・ピーターソン", "book": "『生き抜くための12のルール』", "category": "mindset"},
    {"quote": "心配事の97%は実際には起こらない。", "author": "マーク・トウェイン", "book": "", "category": "mindset"},
    {"quote": "何かを捨てないと、何も始められない。", "author": "スティーブ・ジョブズ", "book": "", "category": "mindset"},
    {"quote": "過去は変えられないが、未来は自分の手の中にある。", "author": "ピーター・ドラッカー", "book": "", "category": "mindset"},

    # ── 成長・学び ──
    {"quote": "成長とは、快適な領域の外に出ることから始まる。", "author": "ブレネー・ブラウン", "book": "『本当の勇気は「弱さ」を認めること』", "category": "growth"},
    {"quote": "学ぶことをやめた者は老いる。学び続ける者はいつまでも若い。", "author": "ヘンリー・フォード", "book": "", "category": "growth"},
    {"quote": "才能は生まれつきではない。努力の積み重ねだ。", "author": "アンジェラ・ダックワース", "book": "『GRIT やり抜く力』", "category": "growth"},
    {"quote": "失敗は終わりではない。それを諦めた時が終わりだ。", "author": "稲盛和夫", "book": "『生き方』", "category": "growth"},
    {"quote": "10,000時間の法則。一流になるには、それだけの時間を捧げる覚悟がいる。", "author": "マルコム・グラッドウェル", "book": "『天才！ 成功する人々の法則』", "category": "growth"},
    {"quote": "フィードバックは成長の朝食だ。", "author": "ケン・ブランチャード", "book": "『1分間マネジャー』", "category": "growth"},
    {"quote": "読書は他人の頭で考えることだ。自分の頭で考える訓練も忘れるな。", "author": "ショーペンハウアー", "book": "『読書について』", "category": "growth"},
    {"quote": "知識に投資することは、常に最大の利益を生む。", "author": "ベンジャミン・フランクリン", "book": "", "category": "growth"},
    {"quote": "弱さを見せられる人間が、一番強い。", "author": "ブレネー・ブラウン", "book": "『Dare to Lead』", "category": "growth"},
    {"quote": "人は教えることで、最もよく学ぶ。", "author": "セネカ", "book": "『道徳書簡集』", "category": "growth"},
    {"quote": "問題は、その問題を作った時と同じ考え方では解決できない。", "author": "アルベルト・アインシュタイン", "book": "", "category": "growth"},

    # ── リーダーシップ・経営 ──
    {"quote": "リーダーとは、道を示し、道を歩み、道を作る人だ。", "author": "ジョン・C・マクスウェル", "book": "『「人の上に立つ」ために本当に大切なこと』", "category": "leadership"},
    {"quote": "まず自分を管理できない者に、他人を管理する資格はない。", "author": "ピーター・ドラッカー", "book": "『経営者の条件』", "category": "leadership"},
    {"quote": "ビジョンなき経営は、羅針盤なき航海と同じだ。", "author": "孫正義", "book": "", "category": "leadership"},
    {"quote": "最高のリーダーは、リーダーがいなくても機能する組織を作る。", "author": "ラオ・ツー（老子）", "book": "『道徳経』", "category": "leadership"},
    {"quote": "人を動かすには、まず相手の欲しがっているものを理解せよ。", "author": "デール・カーネギー", "book": "『人を動かす』", "category": "leadership"},
    {"quote": "戦略とは、何をやらないかを決めることだ。", "author": "マイケル・ポーター", "book": "", "category": "leadership"},
    {"quote": "顧客に聞いたら、もっと速い馬が欲しいと言っただろう。", "author": "ヘンリー・フォード", "book": "", "category": "leadership"},
    {"quote": "まず理解に徹し、そして理解される。", "author": "スティーブン・R・コヴィー", "book": "『7つの習慣』", "category": "leadership"},
    {"quote": "文化は戦略を朝食に食べる。", "author": "ピーター・ドラッカー", "book": "", "category": "leadership"},
    {"quote": "信頼は最速の経営ツールだ。", "author": "スティーブン・M・R・コヴィー", "book": "『スピード・オブ・トラスト』", "category": "leadership"},
    {"quote": "良いアイデアに出会ったら、すぐやれ。明日には誰かがやっている。", "author": "藤田晋", "book": "『渋谷ではたらく社長の告白』", "category": "leadership"},

    # ── 創造性・イノベーション ──
    {"quote": "イノベーションは、リーダーとフォロワーを区別する唯一の基準だ。", "author": "スティーブ・ジョブズ", "book": "", "category": "creativity"},
    {"quote": "創造性とは、つながりを見つけることだ。", "author": "スティーブ・ジョブズ", "book": "", "category": "creativity"},
    {"quote": "制約はクリエイティビティの母である。", "author": "マリッサ・メイヤー", "book": "", "category": "creativity"},
    {"quote": "まだ存在しないものを想像できる人だけが、それを創れる。", "author": "アラン・ケイ", "book": "", "category": "creativity"},
    {"quote": "常識とは、18歳までに身につけた偏見のコレクションだ。", "author": "アルベルト・アインシュタイン", "book": "", "category": "creativity"},
    {"quote": "最高の仕事は、好きなことと得意なことの交差点にある。", "author": "ジム・コリンズ", "book": "『ビジョナリー・カンパニー2』", "category": "creativity"},
    {"quote": "シンプルであることは、複雑であることよりも難しい。", "author": "スティーブ・ジョブズ", "book": "", "category": "creativity"},
    {"quote": "異なる分野の知識を組み合わせる者が、新しい価値を作る。", "author": "フランス・ヨハンソン", "book": "『メディチ・エフェクト』", "category": "creativity"},
    {"quote": "アイデアは実行されなければ、ただの幻だ。", "author": "トーマス・エジソン", "book": "", "category": "creativity"},
    {"quote": "破壊的イノベーションは、いつも業界の外からやってくる。", "author": "クレイトン・クリステンセン", "book": "『イノベーションのジレンマ』", "category": "creativity"},

    # ── 人間関係・コミュニケーション ──
    {"quote": "与えよ、さらば与えられん。", "author": "アダム・グラント", "book": "『GIVE & TAKE』", "category": "mindset"},
    {"quote": "人生の質は、コミュニケーションの質で決まる。", "author": "トニー・ロビンズ", "book": "", "category": "mindset"},
    {"quote": "あなたは、最も多くの時間を共にする5人の平均だ。", "author": "ジム・ローン", "book": "", "category": "growth"},
    {"quote": "聞くことは、最も過小評価されているスキルだ。", "author": "バーナード・フェラーリ", "book": "『パワー・リスニング』", "category": "leadership"},

    # ── お金・ビジネス ──
    {"quote": "金持ちは資産を買い、貧乏人は負債を買う。", "author": "ロバート・キヨサキ", "book": "『金持ち父さん 貧乏父さん』", "category": "action"},
    {"quote": "価格は払うもの、価値は得るものだ。", "author": "ウォーレン・バフェット", "book": "", "category": "mindset"},
    {"quote": "顧客の問題を解決すれば、お金は後からついてくる。", "author": "松下幸之助", "book": "『道をひらく』", "category": "leadership"},

    # ── 日本の経営者 ──
    {"quote": "動機善なりや、私心なかりしか。", "author": "稲盛和夫", "book": "『生き方』", "category": "mindset"},
    {"quote": "素直な心になりましょう。素直な心はあなたを強く正しく聡明にいたします。", "author": "松下幸之助", "book": "『素直な心になるために』", "category": "mindset"},
    {"quote": "夢なき者に理想なし、理想なき者に計画なし、計画なき者に実行なし。", "author": "吉田松陰", "book": "", "category": "action"},
    {"quote": "現状維持は後退の始まりである。", "author": "ウォルト・ディズニー", "book": "", "category": "growth"},
    {"quote": "一日一生。今日という日を精一杯生きる。", "author": "酒井雄哉", "book": "『一日一生』", "category": "mindset"},
]

# カテゴリ名の日本語マッピング
CATEGORY_EMOJI = {
    "action": "🔥 行動",
    "mindset": "🧠 心構え",
    "growth": "🌱 成長",
    "leadership": "👑 リーダーシップ",
    "creativity": "💡 創造性",
}


def load_history():
    """配信履歴を読み込む"""
    if HISTORY_PATH.exists():
        return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    return []


def save_history(history):
    """配信履歴を保存"""
    HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_PATH.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")


def pick_quote():
    """今日の名言を選ぶ（重複回避）"""
    history = load_history()

    # 全名言を使い切ったらリセット
    if len(history) >= len(QUOTES):
        history = []

    # 未使用の名言からランダムに選ぶ
    available = [i for i in range(len(QUOTES)) if i not in history]
    chosen_index = random.choice(available)

    # 履歴に追加
    history.append(chosen_index)
    save_history(history)

    return QUOTES[chosen_index]


def format_discord_message(quote_data):
    """Discord送信用のメッセージを整形"""
    cat_label = CATEGORY_EMOJI.get(quote_data["category"], "📖")
    book_str = f"\n> 📖 {quote_data['book']}" if quote_data["book"] else ""

    weekday_ja = ["月", "火", "水", "木", "金", "土", "日"][TODAY.weekday()]

    message = (
        f"# ✨ 社長に贈る言葉 — {TODAY.strftime('%Y/%m/%d')}（{weekday_ja}）\n"
        f"\n"
        f"> **「{quote_data['quote']}」**\n"
        f"> \n"
        f"> — {quote_data['author']}{book_str}\n"
        f"\n"
        f"_{cat_label}_\n"
        f"\n"
        f"社長、今日も最高の一日にしましょう！ — COO SAEPIN"
    )
    return message


def send_to_discord(message):
    """DiscordのWebhookにメッセージを送信"""
    payload = {"content": message}
    response = requests.post(WEBHOOK_URL, json=payload)
    if response.status_code == 204:
        print("✅ Discord送信成功")
    else:
        print(f"❌ Discord送信失敗: {response.status_code} - {response.text}")


def main():
    print(f"✨ 今日の一言を選択中... ({TODAY.strftime('%Y-%m-%d')})")
    print(f"  名言データベース: {len(QUOTES)}件")

    quote_data = pick_quote()
    message = format_discord_message(quote_data)

    print("\n--- 送信内容 ---")
    print(message)
    print("--- ここまで ---\n")

    send_to_discord(message)

    # LINE にも同時配信
    from line_notify import send_to_line
    send_to_line(message)


if __name__ == "__main__":
    main()
