# NotebookLMに未追加のGeminiメモを一括追加するスクリプト
# 使い方:
#   1. notebooklm login を実行してログイン
#   2. python add_missing_sources.py を実行
#
# 前提: notebooklm use 7620e75a-3a5c-4943-9be1-103348528422 済み

import sys
import os
import subprocess
import time

sys.stdout.reconfigure(encoding='utf-8')

NOTEBOOK_ID = '7620e75a-3a5c-4943-9be1-103348528422'

# 追加するGeminiメモのリスト（Drive File ID, タイトル）
NEW_SOURCES = [
    # 2026.02（未追加分）
    ('18CauXUt-TXEO5TBG1yIW6Ivg9NMq1jnrrHnXYYQxWsw', '2026/02/08 23:34 JST に開始した会議 - Gemini によるメモ'),
    ('1uSSut7f28GSdyZSgQvwBNhPsl5d3hb-YjinuJBBAhkY', '2026/02/23 23:57 JST に開始した会議 - Gemini によるメモ'),
    # 2026.03（3/24以降の未追加分）
    ('1TbGnnWKKPggIRElSc3ipqi94ix3NsGTHMXHjyIyTi-w', '2026/03/24 20:52 - Gemini によるメモ'),
    ('1-XK0dPwGPlkPviWjpgn2Enf_Fniy51K68IMX_PtzLq4', '2026/03/25 20:56 - Gemini によるメモ'),
    ('1VvFj8cRfgOn3_wDshAgYvip1fFVqBgjgTIriU-iqGxs', '2026/03/26 20:52 - Gemini によるメモ'),
    ('1rPjMXuhxM71PYGDHgMPwDjYmBsemc1uCOnLL8zWZxEI', '2026/03/27 20:53 - Gemini によるメモ'),
    ('1IrQJfux2Po2RDJMeJDpF5MRaNWMcl7yMmsdQY1UpfN0', '2026/03/28 10:51 - Gemini によるメモ'),
    ('1UdP7bgZ5rGrxlObOVT90qY-SsjHUyNW_l0FTU8Thp-Q', '2026/03/30 20:56 - Gemini によるメモ'),
    ('1h5WseMr2YaQ5Jrtly8tvDpD0Ppyd-r9KJhJ9w-9q_Yc', '2026/03/31 15:46 - Gemini によるメモ'),
    ('1kj5kCSEuS6LHpLd9g9oAFefAAd1Ig3aXMWLmF38ENo0', '2026/03/31 15:54 - Gemini によるメモ'),
    ('1kRji4tolNZUnGEFjBIFE52v1UkFujvWaERIXQpnOyRE', '2026/03/31 20:22 - Gemini によるメモ'),
    ('1VtybjszoXiY80EoyF0zCvGS1_eoGupX55cHjMi8amrU', '2026/03/31 20:47 - Gemini によるメモ'),
    ('1mgjy6WUwztFi4HAGH116AhncqkilVy6IZeO5xJBSNbU', '2026/03/15 19:51 - Gemini によるメモ'),
    # 2026.04（全部新規）
    ('1_2NDGZX5xsyZp4ICfQSjKsOKJssjRV3cDqHk_TjiWLc', '2026/04/01 17:40 - Gemini によるメモ'),
    ('153OMxT8Qx-PdYE-gvE1ObIirltMox09-tGOudY-fJ74', '2026/04/01 17:49 - Gemini によるメモ'),
    ('1Phfj6L4i0AB6E-2HL9wjFstlFsQ22XYn31E0DCs_LFw', '2026/04/01 19:51 - Gemini によるメモ'),
    ('1tblpjZxkem-7AJzcJkh9pgDmPV2bqrSHA9Pz_oXkD3I', '2026/04/01 20:49 - Gemini によるメモ'),
    ('1HU36kPArBq3CeelAxdCir1YdFMxeWlUE89xZ4qqCzmQ', '実践FB会 - 2026/04/02 00:18 - Gemini によるメモ'),
    ('1yX6Yicf1EN7XKdzZjCJpCOat4dPFPhLnx9vfhZvDE-A', '実践FB会 - 2026/04/02 00:41 - Gemini によるメモ'),
    ('1qqtIH5VcqwvE_dMGK9KGerMupMp72LxuOVQwvRBR0gc', '実践FB会 - 2026/04/02 20:40 - Gemini によるメモ'),
]


def main():
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'

    # ノートブック選択
    print(f'ノートブック {NOTEBOOK_ID} を選択中...')
    subprocess.run(
        ['notebooklm', 'use', NOTEBOOK_ID],
        capture_output=True, text=True, encoding='utf-8', env=env, timeout=30
    )

    added = 0
    errors = 0
    for i, (fid, title) in enumerate(NEW_SOURCES):
        print(f'[{i+1}/{len(NEW_SOURCES)}] 追加中: {title}...', flush=True)
        result = subprocess.run(
            ['notebooklm', 'source', 'add-drive', fid, title],
            capture_output=True, text=True, encoding='utf-8', env=env, timeout=30
        )
        if result.returncode == 0:
            added += 1
            print(f'  → OK')
        else:
            errors += 1
            err = (result.stderr or result.stdout or 'unknown').strip()[:100]
            print(f'  → エラー: {err}')

        # API制限回避のため少し待つ
        time.sleep(2)

    print(f'\n完了: {added}件追加 / {errors}件エラー')


if __name__ == '__main__':
    main()
