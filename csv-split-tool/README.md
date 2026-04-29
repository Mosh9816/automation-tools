# 🌟 CSV分割ツール（CSV Split Tool）
 Excel 形式の元データ（ブロック構造）を読み取り、  
 **項目番号ごとに 30 行 × 15 列のデータを自動で分割し、CSV として出力する Excel VBA ツール**です。 
 
元データのブロック構造を解析し、  
**項目番号（例：123_100）ごとに CSV を自動生成**することで、手作業での切り出し作業を大幅に効率化します。

---

## 📝 概要（Overview）
このツールは、Excel 形式の元データを読み取り、  
**ブロック単位でデータを抽出 → CSV に自動保存**するための Excel マクロです。

- 手作業での切り出し・保存作業を削減
- ブロック構造を自動解析
- 出力フォルダを自動生成
- 実務フローに沿ったシンプルな操作
- ポートフォリオとしても読みやすい構成

元データは以下のような構造を想定しています：
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/203c10be-e6f2-48f1-9aa8-5a10620d38ca" />

---

## 🎯 目的（Purpose）
- ブロック構造のデータを自動で分割し CSV 化
- 手作業によるコピペ作業をゼロに
- 項目番号ごとのファイル作成を自動化
- 実務でよくある「データ分割作業」を効率化

---

## ⚙️ 機能（Features）

### ① 元データ（Excel）を自動読み込み
- デスクトップの 元データ.xlsx を自動で開く
- すでに開いている場合は安全に閉じてから処理開始

### ② ブロック構造（123_100 → 30行）を解析
- `###_###` の形式を項目番号として認識
- その直後の ヘッダー + 30行 を抽出

### ③ CSV ファイルとして自動保存
- ファイル名は以下の形式で出力
```text
データ_105_240430.csv
```

### ④ 出力フォルダを自動生成
- 出力先は デスクトップの「CSV出力」フォルダ  
**※ユーザーが事前に「CSV出力」フォルダを作る必要はありません。(フォルダが存在しない場合は 自動で作成されます)**
```text
C:\Users\<ユーザー名>\Desktop\CSV出力\
```

---

## 🧩 処理フロー（Flow）
1. 日付（yymmdd）を入力  
2. デスクトップの元データ.xlsx を自動で開く 
3. ブロック構造を上から順に解析  
4. 項目番号を検出  
5. ヘッダー + 30行を新規ブックにコピー  
6. CSV として保存  
7. 次のブロックへ進む  
8. 最後まで処理したら完了メッセージ表示  

---

## 💻 使用コード（Main Macro）

```vba
Sub CSV分割()

    Dim 元データ As Workbook
    Dim ws As Worksheet
    Dim 行 As Long
    Dim 最終行 As Long
    Dim raw As String
    Dim num As String
    Dim 日付 As String
    Dim 出力フォルダ As String
    Dim 出力ファイル名 As String
    Dim 新規WB As Workbook
    Dim i As Long, c As Long
    Dim 元データパス As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '---------------------------------------------
    ' ★ 日付入力
    '---------------------------------------------
    日付 = InputBox("出力したい日付を入力してください（例：240430）")
    If 日付 = "" Then GoTo END_PROC

    '---------------------------------------------
    ' ★ 出力フォルダ（ユーザー名を自動取得）
    '---------------------------------------------
    出力フォルダ = Environ("USERPROFILE") & "\Desktop\CSV出力\"

    If Dir(出力フォルダ, vbDirectory) = "" Then
        MkDir 出力フォルダ
    End If

    '---------------------------------------------
    ' ★ 元データ（Excel）を直接パス指定で開く
    '---------------------------------------------
    元データパス = Environ("USERPROFILE") & "\Desktop\元データ.xlsx"

    ' すでに開いていたら閉じる（安全対策）
    On Error Resume Next
    Workbooks("元データ.xlsx").Close SaveChanges:=False
    On Error GoTo 0

    Set 元データ = Workbooks.Open(元データパス)
    Set ws = 元データ.Sheets(1)

    '---------------------------------------------
    ' ★ ブロック構造の解析
    '---------------------------------------------
    最終行 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    行 = 1

    Do While 行 <= 最終行

        raw = ws.Cells(行, 1).Value

        '--- 項目番号（例：123_100）を検出 ---
        If raw Like "###_###" Then

            ' 後半番号だけ抽出（例：123_105 → 105）
            num = Split(raw, "_")(1)

            ' 出力ファイル名
            出力ファイル名 = 出力フォルダ & "データ_" & num & "_" & 日付 & ".csv"

            '-----------------------------------------
            ' ★ 新規ブック作成
            '-----------------------------------------
            Set 新規WB = Workbooks.Add

            ' ヘッダー行
            For c = 1 To 15
                新規WB.Sheets(1).Cells(1, c).Value = ws.Cells(行 + 1, c).Value
            Next c

            ' データ30行
            For i = 1 To 30
                For c = 1 To 15
                    新規WB.Sheets(1).Cells(i + 1, c).Value = ws.Cells(行 + 1 + i, c).Value
                Next c
            Next i

            '-----------------------------------------
            ' ★ CSV 保存
            '-----------------------------------------
            新規WB.SaveAs Filename:=出力ファイル名, FileFormat:=xlCSV
            新規WB.Close SaveChanges:=False

            ' 次のブロックへ（32行進める）
            行 = 行 + 32

        Else
            行 = 行 + 1
        End If

    Loop

    MsgBox "CSV の分割が完了しました！", vbInformation

'---------------------------------------------
' ★ 後処理
'---------------------------------------------
END_PROC:
    On Error Resume Next
    元データ.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub
```
---

## 📌 ポイント（Key Points）
- ブロック構造を自動解析し、項目番号ごとに CSV を生成
- 出力フォルダは自動作成されるため、事前準備が不要
- 元データ.xlsx を自動で閉じる安全設計
- 実務でよくある「データ分割作業」を完全自動化
- シンプルなコードで保守性も高い

---

## 📂 想定利用シーン（Use Cases）
- ブロック構造のデータを扱う業務
- 項目番号ごとにファイルを切り出す作業
- 大量データの CSV 化
- 手作業ミスを減らしたい現場向け

---

## 🌈 今後の拡張案（Future Enhancements）
- 出力ファイル名のカスタマイズ
- ブロック行数の可変対応
- 元データの自動検証（欠損チェックなど）
- UI ボタン化による操作性向上
