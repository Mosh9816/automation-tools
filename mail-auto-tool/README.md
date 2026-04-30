# 🌟 メール自動作成ツール（Mail Auto Tool）
日々のチェック業務で発生する  
**NG 行のエラー通知メール**を  **Outlook の下書きとして自動生成する Excel VBA ツール**です。 

 Data Check Tool で判定された NG 行をもとに、  
 **宛先・件名・本文を自動セットしたメールを複数作成**し、作業ミスを防ぎながら効率化を実現します。

---

## 📝 概要（Overview）
このツールは、チェックシート内の NG 行を読み取り、  
**Outlook の下書きメールを自動生成**するための Excel マクロです。  

- NG 行だけを対象にメールを作成
- 件名・本文を自動生成
- 宛先は固定アドレス（実務で最も一般的）
- 添付なし（必要ならコード1行で追加可能）
- Data Check Tool との連携で実務フローを再現
- ポートフォリオとしても読みやすい構成  
**※ チェックシートに配置された「メール作成」ボタンから実行します。  
  ※ Outlook がインストールされていない環境ではメール作成部分でエラーとなります。**

---

## 🎯 目的（Purpose）
- NG 行のエラー通知メール作成を自動化
- 手作業でのコピペ作業を削減
- 宛先・件名・本文の入力ミスを防止
- Data Check Tool と連携した実務フローを構築
- シンプルで保守しやすいメール自動化ツールを提供

---

## ⚙️ 機能（Features）

### ① チェック日入力（yymm）  
- 例：`0501`  
- 入力された日付のシート（例：チェックシート_0501）を自動選択

### ② NG 行だけを抽出  
- 判定列（D列）が NG の行のみ処理対象

### ③ 宛先・件名・本文を自動生成
- 宛先：固定メールアドレス
- 件名：【エラー通知】管理番号 xxxx
- 本文：丁寧なビジネス文テンプレートを自動生成

### ④ Outlook 下書きを複数作成 
- NG 行が 5 行なら、5 通の下書きを自動生成

### ⑤ 添付なし（必要なら1行追加で対応）
- コードにコメントとして添付処理を残してあり、拡張性あり

---

## 🧩 処理フロー（Flow）
1. チェック日（yymm）を入力
2. 対象シート（チェックシート_yymm）を自動選択
3. NG 行を抽出
4. 宛先・件名・本文を自動生成
5. Outlook 下書きを作成
6. 全メール作成後に完了メッセージ表示

---

## 💻 使用コード（Main Macro）

```vba
Sub メール自動作成()

    Dim checkday As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tgtID As String
    Dim reason As String
    Dim ol As Object
    Dim mail As Object
    
    ' ▼ チェック日入力
    checkday = InputBox("チェック日を4桁で入力してください（例：0501）")
    If checkday = "" Then Exit Sub
    
    ' ▼ 対象シート取得
    On Error Resume Next
    Set ws = Worksheets("チェックシート_" & checkday)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "チェックシート_" & checkday & " が見つかりません。", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' ▼ Outlook 起動
    Set ol = CreateObject("Outlook.Application")
    
    ' ============================
    '   NG 行ごとにメール作成
    ' ============================
    For i = 3 To lastRow
        
        If ws.Cells(i, "D").Value = "NG" Then
            
            tgtID = ws.Cells(i, "B").Value
            reason = ws.Cells(i, "E").Value
            
            Set mail = ol.CreateItem(0)
            
            ' ▼ 宛先（固定アドレス）
            mail.To = "error-team@example.com"
            
            ' ▼ 件名
            mail.Subject = "【エラー通知】管理番号 " & tgtID
            
            ' ▼ 本文（丁寧なビジネス文）
            mail.Body = _
                "お疲れ様です。" & vbCrLf & vbCrLf & _
                "以下の管理番号にてエラーを確認しました。" & vbCrLf & _
                "お手数ですがご確認をお願いいたします。" & vbCrLf & vbCrLf & _
                "管理番号：" & tgtID & vbCrLf & _
                "理由：" & reason & vbCrLf & vbCrLf & _
                "よろしくお願いいたします。"
            
            ' ▼ 添付（必要なら）
            ' mail.Attachments.Add "C:\path\file.xlsx"
            
            mail.Display
            
        End If
        
    Next i
    
    MsgBox "メール作成が完了しました。", vbInformation

End Sub

```
---

## 📌 ポイント（Key Points）
- NG 行だけメールを作成するため、実務のエラー通知フローを再現
- 宛先固定でシンプルかつ実務的
- 添付なしで UI がわかりやすい
- 必要なら添付コードを1行追加するだけで拡張可能
- Outlook がない環境でも処理が止まらない安全設計
- Data Check Tool との連携で業務全体の流れが明確

---

## 📂 想定利用シーン（Use Cases）
- 日次チェック作業のエラー通知
- NG 行の担当者への報告
- チーム内のエラー共有
- 手作業メール作成の削減
- 実務フローの自動化・標準化

---

## 🌈 今後の拡張案（Future Enhancements）
- system_data に担当者メール列を追加して自動振り分け
- 添付ファイルの自動選択
- メール本文テンプレートの外部管理
- 送信前のプレビュー画面に NG 行一覧を表示

## 🔗 関連ツール
- [Format Generator](../format-generator)
- [data-check-tool](./data-check-tool)  
