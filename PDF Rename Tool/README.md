# 🌟 PDF Rename Tool（PDF 自動リネームツール）
PDF の 1 ページ目に記載された文書名（請求書・納品書・契約書など）と  
**system_data の管理番号・キーワード情報を照合し、  
PDF ファイル名を自動でリネームする Excel VBA ツール**です。 

業務でよくある「PDF の命名規則統一」「台帳との突合」「大量ファイルの整理」を  
シンプルな操作で自動化できます。

---

## 📝 概要（Overview）
このツールは、指定フォルダ内の PDF を順番に読み取り、  
**system_data シートの管理番号・文書種別（キーワード）をもとに  
管理番号_文書名.pdf の形式へ自動リネーム**します。    

- フォルダ選択ダイアログで PDF フォルダを指定
- system_data の管理番号・キーワードを参照
- PDF を順番に処理してリネーム
- Acrobat 不要（Edge/Chrome で開ける PDF でOK）
- 実務の「台帳 × PDF 命名統一」フローを再現
- ポートフォリオとしても読みやすい構成

**※ リネームシートに配置された「リネーム開始」ボタンから実行します。**

---

## 🎯 目的（Purpose）
- PDF の命名規則を統一し、管理しやすくする
- system_data（台帳）との突合を自動化
- 手作業でのリネーム作業を削減
- 実務でよくある「大量 PDF の整理」を再現
- シンプルで保守しやすいツールとして提供

---

## ⚙️ 機能（Features）

### ① フォルダ選択ダイアログ  
- ユーザーが PDF の入ったフォルダを自由に選択可能
- ダウンロードフォルダでもデスクトップでもOK

### ② system_data との照合  
- B列：管理番号
- D列：キーワード（文書名）
- 見出しは 2 行目、データは 3 行目から

### ③ PDF の自動リネーム  
- 管理番号_文書名.pdf の形式に統一  
  例：10001_請求書.pdf

### ④ Acrobat 不要  
- PDF の中身を直接解析しないため、無料環境で動作
- Edge/Chrome で開ける PDF で問題なし

---

## 🧩 処理フロー（Flow）
1. system_data シートを準備
2. ユーザーが PDF フォルダを選択
3. system_data の 3 行目から順にデータを読み込む
4. PDF を順番に処理
5. 管理番号_文書名.pdf にリネーム
6. 完了メッセージを表示

---

## 💻 使用コード（Main Macro）

```vba
Sub PDFリネーム()

    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim mngNo As String
    Dim docName As String
    Dim newName As String
    
    ' ▼ system_data シート
    Set ws = ThisWorkbook.Worksheets("system_data")
    
    ' ▼ 管理番号の最終行（B列）
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' ▼ フォルダ選択ダイアログ
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "PDF が入っているフォルダを選択してください"
        If .Show <> -1 Then Exit Sub   ' キャンセル時は終了
        folderPath = .SelectedItems(1)
    End With
    
    ' ▼ FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' ▼ データ開始行（見出しが2行目 → データは3行目）
    rowIndex = 3
    
    ' ============================
    '   PDF ファイルを順番に処理
    ' ============================
    For Each file In folder.Files
        
        If LCase(fso.GetExtensionName(file.Name)) = "pdf" Then
            
            If rowIndex > lastRow Then Exit For
            
            ' ▼ 管理番号（B列）・文書名（D列）を取得
            mngNo = CStr(ws.Cells(rowIndex, "B").Value)
            docName = CStr(ws.Cells(rowIndex, "D").Value)
            
            ' ▼ 新しいファイル名：管理番号_文書名.pdf
            newName = mngNo & "_" & docName & ".pdf"
            
            ' ▼ リネーム実行
            file.Name = newName
            
            rowIndex = rowIndex + 1
            
        End If
        
    Next file
    
    MsgBox "PDF のリネームが完了しました。", vbInformation

End Sub

```
---

## 📌 ポイント（Key Points）
- system_data の並び順と PDF の並び順を対応させてリネーム
- 実務の「台帳 × PDF 命名統一」フローを再現
- フォルダ選択ダイアログで操作性が高い
- Acrobat 不要で環境依存が少ない
- コードがシンプルで保守しやすい
- 他ツールと同じ構成でポートフォリオに統一感がある

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
- PDF の中身を検索して管理番号を抽出する高度版

## 🔗 関連ツール
- [Format Generator](../format-generator)  
  チェック作業で使用する system_data を生成するツール。

- [mail-auto-tool](../mail-auto-tool)  
  system_data を利用して NG 行のエラー通知メールを自動作成するツール。
