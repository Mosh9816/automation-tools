# 🌟 PDF Rename Runner（PDF 自動リネームフロー）
GitHub から ZIP を取得し、**解凍 → Excel マクロで PDF を自動リネーム →バックアップ作成**までを  
一括で実行する Power Automate Desktop（PAD）フローです。

実務でよくある「PDF の命名規則統一」「台帳との突合」「大量ファイルの整理」を  
**ブラウザ操作 × Excel マクロ × PAD** の組み合わせで再現しています。

---

## 📝 概要（Overview）
このフローは、GitHub 上のサンプル PDF セットをダウンロードし、    
Excel マクロ（PDFリネームツール.xlsm）をユーザーが選択、  
PDF を自動リネームしてバックアップ作成までを自動化する仕組みです。

- ZIP ダウンロード → 解凍
- Excel マクロで PDF を一括リネーム
- リネーム済みフォルダの整理
- バックアップフォルダへコピー
- 完了メッセージ表示
  
**※PAD × Excel の連携で、実務の「台帳 × PDF 命名統一」フローを再現しています。**

---

## 🎯 目的（Purpose）
- PDF の命名規則を自動で統一
- Excel 台帳（system_data）との突合を自動化
- 手作業のリネーム作業を削減
- ダウンロード → 整理 → バックアップまでの一連処理を自動化
- 実務でよくある「大量 PDF の整理フロー」をポートフォリオとして再現

---

## 📂 リポジトリ構成（Repository Structure）
```text
pdf-rename-runner/
 ├─ PDFリネームツール.xlsm     ← リネーム処理を行う Excel マクロ
 ├─ dummy-pdf-set.zip           ← サンプル PDF セット
 ├─ リネームフロー.txt          ← PAD に貼り付けて使うフローのテキスト版
 └─ README.md
```
✔ **PAD フローはテキスト貼り付け方式**    
   環境依存が少なく、ポートフォリオとしても読みやすい構成です。

---

## ⚠️ Excel マクロの選択について（重要）
このフローでは、Excel マクロ（PDFリネームツール.xlsm）の場所を**ユーザーがダイアログボックスで選択**できます。

### ① Excel マクロ選択ダイアログ（PAD 側）
  - フロー開始後、次のダイアログが表示されます：  
    ここでは、事前にダウンロードした**PDFリネームツール.xlsm** を選択してください。
  <img width="927" height="575" alt="image" src="https://github.com/user-attachments/assets/5775c6fa-7030-47d7-bfe0-fa9ac36337ee" />

### ② フォルダ選択ダイアログ（Excel マクロ側）
  - Excel マクロ実行中に、次のダイアログが表示されます：  
    ここでは**ZIP 解凍によって作成された dummy-pdf-set フォルダ**を選択してください。
  <img width="931" height="571" alt="image" src="https://github.com/user-attachments/assets/ffcf9749-d6ae-4e4f-bf94-e764b2f691aa" />

---

## ⚙️ 機能（Features）
### ① ZIP の自動ダウンロード  
- GitHub の ZIP を Edge で開き、「Download raw file」をクリック
- ダウンロード完了まで待機（0KB 問題を回避）

### ② ZIP の自動解凍
- %USERPROFILE%\Downloads に展開
- フォルダ構成を自動で整備

### ③ Excel マクロで PDF を一括リネーム
- ユーザーが選択した Excel マクロを起動
- system_data の管理番号・文書名を参照
- 管理番号_文書名.pdf の形式に統一
- 実務の「台帳 × PDF 命名統一」フローを再現

### ④ フォルダ名の変更
- dummy-pdf-set → リネーム完了 に自動変更

### ⑤ バックアップ作成
- バックアップ フォルダを自動生成
- リネーム済み PDF をすべてコピー

### ⑥ 完了メッセージ
- 「フローが完了しました。」を表示

---

## 🧩 処理フロー（Flow）
1. GitHub から ZIP をダウンロード
2. ZIP が生成されるまで待機
3. ZIP を解凍
4. ユーザーが Excel マクロを選択
5. PDF を自動リネーム
6. フォルダ名を「リネーム完了」に変更
7. バックアップフォルダを作成
8. PDF をコピー
9. 完了メッセージを表示

---

## ▶️ 実行方法（How to Run）
1. Power Automate Desktop を起動
2. 新規フローを作成
3. リネームフロー.txt を開き、内容をすべてコピー
4. PAD のフロー編集画面に貼り付け
5. 実行時に表示されるダイアログで PDFリネームツール.xlsm を選択する
6. マクロ実行中のダイアログで dummy-pdf-set フォルダを選択する
7. フローを実行するだけで完了

---

## 📌 ポイント（Key Points）
- Excel マクロの場所をユーザーが選択できる柔軟設計
- PAD × Excel の連携で実務フローを再現
- ZIP 完成待機で安定動作
- テキスト版フローで読みやすく、ポートフォリオ向け
- コード・フローともにシンプルで保守しやすい

---

## 📂 想定利用シーン（Use Cases）
- 請求書・納品書などの PDF 整理
- 台帳との突合作業の自動化
- 大量 PDF の命名規則統一
- 業務フローの標準化
- ポートフォリオでの実務再現

---

## 🌈 今後の拡張案（Future Enhancements）
- PDF の 1 ページ目から文書名を自動抽出（上位版）
- system_data に担当者列を追加して振り分け
- リネーム後のログ出力
- PDF 内容から管理番号を抽出する高度版

---

## 🔗 関連ツール（Related Tools）
- [pdf-rename-tool](..//pdf-rename-tool)
