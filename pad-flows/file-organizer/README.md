# 📦 File Organizer（PAD）
Web から取得し、PDF Rename Tool でリネームされた PDF を  
日付フォルダへ自動整理し、バックアップも作成する PAD フローです。

実務でよくある  
**リネーム済み PDF を日付ごとに整理 → バックアップ**という作業を自動化します。

---

## 🎯 目的（Purpose）
- PDF Rename Tool でリネームされた PDF を自動で整理
- 日付フォルダ（例：2026-05-01）を自動生成
- 指定フォルダへ移動し、バックアップも作成
- 手作業のドラッグ＆ドロップを完全に排除

---

## ⚙️ フロー構成（Flow）
### 1. リネーム済み PDF のフォルダを取得  
   例：C:\Users\ユーザー名\Downloads\pdf-set-renamed\

### 2. 今日の日付フォルダを自動生成  
   例：C:\automation-tools\output\2026-05-01\

### 3. PDF を日付フォルダへ移動

### 4. バックアップフォルダを作成  
   例：C:\automation-tools\backup\2026-05-01\

### 5. PDF をバックアップへコピー

### 6. 完了メッセージを表示

---

## 📂 フォルダ構成（Folder Structure）
```text
file-organizer/
├─ flow.json
└─ README.md
```

---

## 🧩 連携フロー（Integration Flow）
このフローは、以下の PAD フローと連携して動作します。  

- web-pdf-downloader  
  → Web から ZIP をダウンロードして解凍

- pdf-rename-runner  
  → PDF Rename Tool を実行してリネーム

- file-organizer（本フロー）  
  → リネーム済み PDF を整理・バックアップ

この 3 フローで**Web → PDF → Rename → 整理** という実務フローが完全に自動化されます。

---


## 📝 使用例（Example）
- web-pdf-downloader で ZIP を取得
- pdf-rename-runner で PDF をリネーム
- file-organizer を実行
- PDF が自動で日付フォルダに整理され、バックアップも作成される

---


## 🔗 関連ツール（Related Tools）
- PDF Rename Tool（Excel VBA）
- web-pdf-downloader（PAD）
- pdf-rename-runner（PAD）
