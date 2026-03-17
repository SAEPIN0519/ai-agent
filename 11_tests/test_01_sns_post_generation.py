"""
テスト01: SNS投稿文の自動生成
対象: 店舗情報・メニュー・イベントからInstagram/X投稿文を生成する機能
"""
import pytest


def generate_sns_post(store_name: str, content: str, platform: str = "instagram") -> str:
    """SNS投稿文を生成するダミー関数（本番はAI APIを呼ぶ）"""
    if platform == "instagram":
        return f"【{store_name}】\n{content}\n\n#カフェ #喫茶店"
    elif platform == "x":
        return f"【{store_name}】{content}"
    return ""


# テスト1: Instagram投稿文が生成される
def test_instagram_post_is_generated():
    result = generate_sns_post("カフェ山田", "本日のランチは日替わりカレーです！")
    assert result != ""
    assert "カフェ山田" in result


# テスト2: X(Twitter)投稿文が生成される
def test_x_post_is_generated():
    result = generate_sns_post("カフェ山田", "今週末はライブイベント開催！", platform="x")
    assert "カフェ山田" in result
    assert "ライブイベント" in result


# テスト3: Instagramのハッシュタグが含まれる
def test_instagram_includes_hashtag():
    result = generate_sns_post("テスト店", "テスト内容", platform="instagram")
    assert "#" in result


# テスト4: 店舗名が必ず投稿文に含まれる
def test_store_name_always_included():
    store_name = "ブルーマウンテン珈琲"
    result = generate_sns_post(store_name, "新メニュー登場")
    assert store_name in result


# テスト5: 不明なプラットフォームは空文字を返す
def test_unknown_platform_returns_empty():
    result = generate_sns_post("テスト店", "内容", platform="unknown")
    assert result == ""
