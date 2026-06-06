# ペット供養祭 申込フォーム → スライドショー パイプライン (GAS)

フォーム送信 → 回答シート → Google スライド（当日のスライドショー）を自動生成する Google Apps Script。

- フォーム: <https://forms.gle/MnTctVaquHdpjLw16>
- 回答スプレッドシート ID: `1EWZeJxLKgoXCZzpBsQ2ny4HxTEbm6P6ucm6awkzYt4A`
- 出力スライド ID: `1EVy_prP1oQ4E662ZqHN4uVOGHp4cNFnr5sxLaHIyTZ4`
- scriptId: `1wkdKD7vCq13W07xNgchSOQ1aphvtI3d0YzyHkwXBbT2Ci619IPof2kaQ`

## このディレクトリの中身

| ファイル | 役割 |
| --- | --- |
| `Code.gs` | パイプライン本体（トリガー・スライド生成・診断・テスト） |
| `appsscript.json` | マニフェスト（タイムゾーン / OAuth スコープ） |
| `.clasp.json` | clasp 用設定（scriptId） |

## デプロイ手順（clasp）

```bash
npm install -g @google/clasp
clasp login
cd gas
clasp push        # ローカル → GASプロジェクトへ反映
# 既存コードを取得したい場合は: clasp pull
```

> `.clasp.json` の `scriptId` は本プロジェクトを指している。`clasp push` で上書きされるため、
> 先に `clasp pull` で現行コードを退避・比較してから push すると安全。

## 初回セットアップ（GASエディタで手動実行）

1. `installFormSubmitTrigger` を実行 → `onFormSubmit` トリガーを設置。
2. `diagnoseLatestResponse` を一度実行 → **DriveApp の認可ダイアログを承認**。
   （DriveApp を使う関数を手動実行して認可を通さないと、自動トリガーから写真にアクセスできない）

## 検証フロー

1. フォームからテスト送信を1件。
2. `diagnoseLatestResponse` を実行し、実行ログで以下を確認:
   - 列構成（ヘッダー）と列解決結果（写真/お名前/ふりがな/参加回/メッセージの idx が全て >= 0）。
   - 参加回ラベルが `10:00 / 12:30 / 15:00 / 16:00` で正しく変換される（旧 12:00/14:00 残存なし）。
   - 写真の Drive ファイルID抽出 → `DriveApp OK` が出る。
3. `testGenerateSlideFromLatestRow` で最新行からスライドを生成し、出力スライドを目視確認。
4. フォーム送信せずレイアウトだけ確認したいときは:
   - `testGenerateSlideWithDummy`（写真なし＋長文メッセージ）
   - `testGenerateSlideEmptyMessage`（メッセージ空欄 → メッセージ枠が出ない）

## 設計上のポイント

- **列参照はヘッダー名（部分一致）ベース**（`COLUMN_KEYS`）。設問の増減・並べ替えに強い。
  文言が変わったら `COLUMN_KEYS` のキーワード候補を追記するだけで追従できる。
- **写真は1送信=1枚（単一URL）前提**。複数連結（`__` や `, `）が来ても先頭1件を採用。
- **参加回ラベル**は `SESSION_RULES` で時刻文字列をマッチさせて変換。
  選択肢を変えたらここを更新。未知の値はログに警告を出して生値をそのまま表示。
- **一言メッセージ**は本文枠として配置。空欄なら枠を出さない。長文は `autoshrinkText` で簡易自動縮小。
- スプレッドシート参照は `openById`（コンテナバインド/スタンドアロン両対応、`getActiveSheet` は不使用）。

## 音楽 (BGM) — 重要な制約

**Apps Script / Slides API には音声を挿入・制御するメソッドが無い。** プログラムからは追加できないため、
スライド側は手動設定で対応する。

推奨手順（最小工数）:

1. BGM 音源を Drive にアップロード。
2. Slides UI で **1枚目のスライド**を開く → メニュー **挿入 > 音声** → Drive の音源を選択。
3. 音声アイコンを選択し「書式設定オプション」で:
   - **再生 = 自動**
   - **スライドショーで音声アイコンを非表示 = ON**（任意）
   - **スライドの切り替え時に停止 = OFF**（全体で流すため）
   - **ループ = ON**（必要なら）

1枚目に置けばスライドショー全体で流れる。音源は後日指定 → 上記2〜3を実施して完了。
