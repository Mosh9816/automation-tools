
📥 Web PDF Downloader（PAD）
Web 上に公開された PDF 一式（ZIP）を自動ダウンロードし、
解凍して PDF Rename Tool に渡すための PAD フローです。

実務でよくある
「資料 ZIP をダウンロード → 展開 → リネーム → 整理」  
という一連の作業を自動化します。

📝 概要（Overview）
GitHub に置いた dummy-pdf-set.zip を自動ダウンロード

ZIP を解凍して PDF を展開

PDF Rename Tool に渡せる状態に整える

file-organizer フローと連携可能

🎯 目的（Purpose）
Web からの資料取得を自動化

PDF Rename Tool との連携を前提とした前処理

ZIP 展開 → PDF 整理の手作業を削減

RPA × Excel VBA の実務フローを再現

⚙️ フロー構成（Flow）
ブラウザ（Edge/Chrome）を起動

GitHub Raw の ZIP URL を開く

Ctrl+S で ZIP をダウンロード

ダウンロードフォルダに保存

ZIP を解凍

展開された PDF フォルダを取得

完了メッセージ表示

※ この後、pdf-rename-runner フローで PDF Rename Tool を実行します。

📂 フォルダ構成（Folder Structure）
コード
web-pdf-downloader/
├─ flow.json
├─ dummy-pdf-set.zip
└─ README.md
📦 dummy-pdf-set.zip の内容
コード
10001_請求書.pdf
10002_納品書.pdf
10003_契約書.pdf
10004_受領書.pdf
10005_発注書.pdf
10006_納品書.pdf
10007_請求書.pdf
10008_契約書.pdf
10009_発注書.pdf
10010_受領書.pdf
system_data（10001〜10010）と一致するため、
PDF Rename Tool で正しくリネームできます。

🔗 関連フロー（Related Flows）
pdf-rename-runner  
→ PDF Rename Tool を PAD から実行するフロー

file-organizer  
→ リネーム後の PDF を整理・移動するフロー
