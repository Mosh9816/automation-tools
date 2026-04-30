# 🌟 データチェックツール（Data Check Tool）
営業日に行うチェック業務での使用を想定して設計した  
**取込データと system_data を照合し、OK/NG を自動判定する Excel VBA ツール**です。 

 取込データ管理番号をもとに  **重複・存在チェック・形式チェック・キーワード照合・システム番号転記**を自動で行い、  
 作業ミスを防ぎながら効率化を実現します。

---

## 📝 概要（Overview）
このツールは、日々のチェック作業で使用するデータを  
**自動で照合・判定・色付け**するための Excel マクロです。  

- 手作業での照合作業を自動化
- NG 行を自動で赤色表示
- システム管理番号を自動転記
- 実務フローに沿ったシンプルな UI
- ポートフォリオとしても読みやすい構成
**※ チェックシートに配置された「チェック開始」ボタンから実行します。**

---

## 🎯 目的（Purpose）
- 取込データのチェック作業を自動化
- 作業者の負担を軽減し、作業スピードを向上
- 重複・入力ミス・照合漏れを防止
- 実務でよくある「データ照合作業」を効率化

---

## ⚙️ 機能（Features）

### ① チェック日入力（yymm）  
- 例：`0430`  
- 入力された日付のシート（例：チェックシート_0501）を自動選択

### ② system_data との照合  
- 管理番号一致チェック
- キーワード有無チェック
- システム管理番号の自動転記

### ③ 重複チェック  
- 同じ管理番号が複数行ある場合は NG 判定

### ④ 形式チェック  
- 数字以外の文字が含まれていれば NG

### ⑤ NG 行の色付け  
- NG 行は 行全体を赤色で強調

### ⑥ 見出しは保護  
- 2行目の見出しは絶対に消さない設計
---

## 🧩 処理フロー（Flow）
1. チェック日（yymm）を入力
2. 対象シート（チェックシート_yymm）を自動選択
3. 初期化（3行目以降の判定・理由・色をクリア）
4. 取込データ管理番号を1行ずつチェック
5. system_data と照合
6. 重複・形式チェック
7. OK/NG 判定
8. NG 行は赤色で強調表示
9. 完了メッセージ表示

---

## 💻 使用コード（Main Macro）

```vba
Sub データチェック実行()

    Dim checkday As String
    Dim ws As Worksheet
    Dim sys As Worksheet
    Dim lastRow As Long
    Dim sysLast As Long
    Dim sysStart As Long
    Dim i As Long, j As Long
    Dim tgtID As String
    Dim found As Boolean
    Dim reason As String
    
    ' ▼ チェック日を入力（例：0501）
    checkday = InputBox("チェック日を半角数字4桁で入力してください（例：0501）")
    If checkday = "" Then Exit Sub
    
    ' ▼ チェック対象シートを自動選択
    On Error Resume Next
    Set ws = Worksheets("チェックシート_" & checkday)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "チェックシート_" & checkday & " が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' ▼ system_data シート
    Set sys = Worksheets("system_data")
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    sysLast = sys.Cells(sys.Rows.Count, "B").End(xlUp).Row
    
    sysStart = 3 ' system_data のデータ開始行
    
    ' ▼ 初期化（3行目以降のみ）
    ws.Range("C3:E" & lastRow).ClearContents
    ws.Range("B3:E" & lastRow).Interior.ColorIndex = xlNone
    
    ' ============================
    '   メインチェック処理
    ' ============================
    For i = 3 To lastRow
        
        tgtID = ws.Cells(i, "B").Value
        reason = ""
        found = False
        
        ' --- 空欄チェック ---
        If tgtID = "" Then
            reason = reason & " / 管理番号が空欄"
        End If
        
        ' --- 重複チェック ---
        For j = 3 To lastRow
            If i <> j And ws.Cells(j, "B").Value = tgtID And tgtID <> "" Then
                reason = reason & " / 重複"
            End If
        Next j
        
        ' --- system_data 照合 ---
        For j = sysStart To sysLast
            If sys.Cells(j, "B").Value = tgtID Then
                found = True
                
                If sys.Cells(j, "D").Value = "" Then
                    reason = reason & " / キーワードなし"
                Else
                    ws.Cells(i, "C").Value = sys.Cells(j, "E").Value
                End If
                
                Exit For
            End If
        Next j
        
        ' --- system_data に存在しない ---
        If Not found Then
            reason = reason & " / system_dataに存在しない"
        End If
        
        ' --- 形式チェック ---
        If tgtID <> "" Then
            If Not IsNumeric(tgtID) Then
                reason = reason & " / 数字以外の文字を含む"
            End If
        End If
        
        ' --- 判定 ---
        If reason = "" Then
            ws.Cells(i, "D").Value = "OK"
        Else
            ws.Cells(i, "D").Value = "NG"
            ws.Cells(i, "E").Value = Mid(reason, 4)
            ws.Rows(i).Interior.Color = RGB(255, 200, 200)
        End If
        
    Next i
    
    MsgBox "チェック完了しました。", vbInformation

End Sub

```
---

## 📌 ポイント（Key Points）
- NG 行を赤色で強調し、視認性を向上
- 見出し（2行目）は絶対に消さない安全設計
- system_data のキーワード照合で実務再現性が高い
- 日付別のチェックシートと連携しやすい
- シンプルな UI とコードで保守性も高い 

---

## 📂 想定利用シーン（Use Cases）
- 日次の取込データチェック
- システム照合作業の効率化
- 重複・入力ミスの検出
- チームでのチェック作業標準化
- 手作業ミスを減らしたい現場向け

---

## 🌈 今後の拡張案（Future Enhancements）
- system_data の自動更新
- NG の種類別に色分け
- チェック結果の自動メール送信
- チェック対象の列を可変にする設定画面の追加

## 🔗 関連ツール
- [Format Generator](../format-generator)
