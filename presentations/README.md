# プレゼンテーション成果物

## 再生成

Python 3.10 以上で、リポジトリのルートから実行します。

```bash
python -m pip install -r presentations/requirements.txt
python scripts/generate_presentation.py
```

スクリプトは `docs/00_INDEX.md` と `docs` 配下の52アカウント文書を読み、次を生成・検証します。

- `M365_Copilot_License_Expansion_ISD_Strategy_20260916.pptx`
- `account_evidence_matrix.csv`（UTF-8 BOM）
- `source_manifest.json`

検証はZIP/XML妥当性、PowerPointとしての再読込、75スライド、52アカウント、未解決プレースホルダーの不在です。実行環境にLibreOfficeがない場合、視覚レンダリング検証は実施できません。レイアウトは16:9、本文11pt以上、短い根拠抜粋と固定領域で生成します。

## 利用上の注意

台帳の有償・無償・合計は活動ユーザーや新規受注ではありません。製品権利、価格、包装、契約条件は営業が確認してください。資料の欠損・矛盾・`[...]` は補完していません。
