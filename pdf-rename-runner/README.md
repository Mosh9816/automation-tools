# 📄 PDF Rename Runner（PDFリネームフロー）
Power Automate Desktop（PAD）で動作する PDF リネーム自動化フローです。  
**GitHub から ZIP をダウンロードし、解凍 → Excel マクロでリネーム → バックアップ作成**までを自動で実行します。

---

## 📌 概要
このフローは、PDF ファイルのリネーム作業を自動化するための Power Automate Desktop フローです。  
GitHub からサンプル PDF セットを取得し、Excel マクロを使ってファイル名を一括変換し、  
バックアップ作成までを含めた一連の処理を自動化します。

---

## 📂 リポジトリ構成
```texr
pdf-rename-runner/
 ├─ PDFリネームツール.xlsm     ← リネーム処理を行う Excel マクロ
 ├─ dummy-pdf-set.zip           ← サンプル PDF セット
 ├─ リネームフロー.txt          ← PAD に貼り付けて使うフローのテキスト版
 └─ README.md
```

- PDFリネームツール.xlsm：Excel マクロ本体
- dummy-pdf-set.zip：動作確認用の PDF セット
- リネームフロー.txt：PAD にそのままコピーして使うフローのテキスト版

**※ リネームフロー.txtを直接 PAD に貼り付けて利用する構成です**

---

## ⚠️ 重要：リネームツール（Excel マクロ）の配置について
このフローでは、以下の Excel マクロファイルを使用します：

```text
PDFリネームツール.xlsm
```

**✔ 必ず ユーザーの Downloads フォルダに配置してください**  
⇒PAD フロー内で以下のパスを参照しているためです：  

```text
%USERPROFILE%\Downloads\PDFリネームツール.xlsm
```

**※Downloads 以外に置くと、  Excel 起動アクションでファイルが見つからずエラーになります。**

---

## ⚙️ フローの処理内容（Step-by-step）
**1. GitHub から ZIP をダウンロード**    
Microsoft Edge を起動し、dummy-pdf-set.zip の「Download raw file」ボタンをクリック。

**2. ZIP が生成されるまで待機**  
File.WaitForFile.Created を使用し、%USERPROFILE%\Downloads\dummy-pdf-set.zip が生成されるまで待機。

**3. ZIP を解凍**  
Downloads フォルダに ZIP を展開。

**4. Excel マクロを実行**  
PDFリネームツール.xlsm を起動し、PDFリネーム マクロを実行。

**5. フォルダ名を変更**   
dummy-pdf-set → リネーム完了 に変更。

**6. バックアップフォルダを作成**  
Downloads 内に バックアップ フォルダを作成。

**7. PDF をバックアップへコピー**  
リネーム済み PDF をすべてバックアップへコピー。

**8. 完了メッセージを表示**  
「フローが完了しました。」と通知。

---

## ▶️ 実行方法（テキスト版を PAD に貼り付けて使用）
1. Power Automate Desktop を起動  
2. 新規フローを作成  
3. リネームフロー.txt を開き、内容をすべてコピー  
4. PAD のフロー編集画面に貼り付ける  
5. PDFリネームツール.xlsm を Downloads フォルダに配置  
6. フローを実行するだけで動作します  

## 📝 補足（テキスト版について）
- リネームフロー.txt は PAD にそのまま貼り付けて使える形式です
- flow.json を使わない構成のため、環境依存が少なく扱いやすい
- GitHub 上で処理内容を確認したい場合にも便利です
