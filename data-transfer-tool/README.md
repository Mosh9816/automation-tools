# 🌟 データ転記ツール（Data Transfer Tool）
日次で更新される元データを安全に転記し、  
**更新日チェック → 転記 → 件数チェック → 保存**を自動で行う
Excel VBA の **日次処理自動化ツール**です。

「クリックで転記開始」ボタンを押すだけで、
実務フローに沿った処理が順番に実行されます。

---

## 📝 概要（Overview）
このツールは、毎日更新される Excel データを
**安全に転記し、件数整合性を確認し、日付付きで保存する** ための VBA マクロです。

- 更新日の鮮度チェック
- A列データの自動転記
- 転記元・転記先の件数チェック
- 日付フォルダの自動生成
- 保存用データ_yyyymmdd.xlsx 形式で保存
- 実務で行っていた日次処理を、個人情報を含まない形で再現しています。

---

## 🎯 目的（Purpose）
- 日次の転記作業を自動化し、作業時間を削減
- 更新日チェックにより誤転記を防止
- 件数整合性チェックで品質を担保
- 日付別保存により履歴管理を容易にする

---

## ⚙️ 機能（Features）
### ① 更新日チェック（鮮度確認）
- 元データの更新日を取得
- 当日更新されていない場合は警告
- 古いデータの誤転記を防止

### ② データ転記（A列コピー）
- A列の最終行まで自動取得
- 新規ブックに値貼り付けで転記
- 書式を持ち込まない安全設計

### ③ 件数チェック（整合性確認）
- 転記元と転記先の件数を比較
- 抽出漏れ・転記漏れを検出
- NG の場合は警告表示

### ④ 保存処理（フォルダ自動作成）
- 保存先フォルダが無ければ自動作成
- 保存用データ_yyyymmdd.xlsx の形式で保存
- 日次履歴を自動で蓄積

---

## 🧩 処理フロー（Flow）
1. 更新日を取得し、当日更新かチェック
2. A列の最終行まで自動取得
3. 新規ブックを生成し、A列を転記
4. 転記元・転記先の件数を比較
5. 保存先フォルダを自動作成
6. 保存用データ_yyyymmdd.xlsx 形式で保存
7. 完了メッセージ表示

---

## 💻 使用コード（Main Macro）
```vba
Sub データ転記開始()

    Dim src As Worksheet
    Dim lastRow As Long
    Dim newWb As Workbook
    Dim savePath As String
    Dim todayStr As String
    Dim srcCount As Long, dstCount As Long
    
    ' ▼ 元データシート
    Set src = Worksheets("元データ")
    
    ' ▼ 更新日チェック
    If Format(FileDateTime(ThisWorkbook.FullName), "yyyymmdd") <> Format(Date, "yyyymmdd") Then
        MsgBox "元データが本日更新されていません。", vbExclamation
        Exit Sub
    End If
    
    ' ▼ 最終行取得
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).Row
    srcCount = WorksheetFunction.CountA(src.Range("A1:A" & lastRow))
    
    ' ▼ 新規ブック作成
    Set newWb = Workbooks.Add
    src.Range("A1:A" & lastRow).Copy
    newWb.Sheets(1).Range("A1").PasteSpecial xlPasteValues
    
    ' ▼ 件数チェック
    dstCount = WorksheetFunction.CountA(newWb.Sheets(1).Range("A:A"))
    
    If srcCount <> dstCount Then
        MsgBox "件数が一致しません。転記を中止します。", vbCritical
        newWb.Close False
        Exit Sub
    End If
    
    ' ▼ 保存処理
    todayStr = Format(Date, "yyyymmdd")
    savePath = "C:\Users\YourName\Desktop\日時処理データ\"
    
    If Dir(savePath, vbDirectory) = "" Then MkDir savePath
    
    newWb.SaveAs savePath & "保存用データ_" & todayStr & ".xlsx"
    newWb.Close False
    
    MsgBox "転記が完了しました。", vbInformation

End Sub
```

---

## 📁 保存先（Save Location）
保存先フォルダ（各自の PC に合わせて変更）：

```text
C:\Users\YourName\Desktop\日時処理データ\
例（ユーザー名が siori の場合）：
```
```text
C:\Users\siori\Desktop\日時処理データ\
保存されるファイル名：
```
```text
保存用データ_20240428.xlsx
```

---

## 🖥 UI（画面イメージ）
- 大きな「クリックで転記開始」ボタン
- 下部に処理ステータスを表示
- 初めて使う人でも迷わないシンプル設計

---

## 📌 想定利用シーン（Use Cases）
- 日次のデータ転記作業
- 履歴を残しながらのデータ管理
- 転記ミス防止が求められる業務
- 手作業のコピー＆ペーストを自動化したい現場

---

🌈 今後の拡張案（Future Enhancements）
- 転記対象列の可変化
- 保存先フォルダの設定画面追加
- 転記結果の自動メール送信
- 転記ログの自動生成
