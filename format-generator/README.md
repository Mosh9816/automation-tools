# 🌟 作業フォーマット自動生成ツール（Format Generator）

実務で使用する **チェックシート（yymm）を自動生成する Excel VBA ツール**です。  
原本シートを元に複数枚のチェックシートを作成し、  
**重複日付のシートは自動スキップ**することで、作業ミスを防ぎながら効率化を実現します。

---

## 📝 概要（Overview）
このツールは、日々のチェック作業で使用するフォーマットを  
**自動で複製・命名・整理**するための Excel マクロです。

- 手作業でのコピー＆リネーム作業を削減  
- 重複作成によるミスを防止  
- 実務フローに沿ったシンプルな UI  
- ポートフォリオとしても読みやすい構成

**※ 原本シートに配置された「シート作成」ボタンから実行します。**

---

## 🎯 目的（Purpose）
- 日次チェック作業のフォーマット作成を自動化  
- 作業者の負担を軽減し、作業スピードを向上  
- シート名の重複によるミスを防止  
- 実務でよくある「フォーマット作成作業」を効率化

---

## ⚙️ 機能（Features）

### ① 日付入力（yymm）
- 例：`0430`  
- 4桁以外はエラー表示

### ② 原本シートをコピー
- 原本シートを最後尾に複製

### ③ シート名を自動設定
- `チェックシート_yymm` の形式で命名

### ④ コピーされた作成ボタンを削除
- 原本の UI をそのままコピーしないように調整

### ⑤ 重複日付のシートはスキップ
- 既に存在する場合はメッセージ表示  
- そのシートだけ削除して次のループへ進む  
- 全体処理は止まらない

---

## 🧩 処理フロー（Flow）
1. 作成枚数を入力  
2. 1枚ずつ日付を入力  
3. 原本シートをコピー  
4. シート名を設定  
5. 重複チェック  
6. 重複ならスキップ  
7. 問題なければ作成完了

---

## 💻 使用コード（Main Macro）

```vba
Sub 作業フォーマット自動生成()

    Dim 入力日 As String
    Dim 原本 As Worksheet
    Dim シート名 As String
    Dim 新シート As Worksheet
    Dim 回数 As Long
    Dim i As Long

    回数 = InputBox("作成するシート数を入力してください（例：3）")
    If 回数 <= 0 Then Exit Sub

    Set 原本 = ThisWorkbook.Sheets("原本")

    For i = 1 To 回数

        入力日 = InputBox("チェック日を4桁で入力してください（例：0430）" & vbCrLf & _
                          "（" & i & "枚目）")

        If 入力日 = "" Then Exit Sub
        If Not 入力日 Like "####" Then
            MsgBox "4桁の数字で入力してください。", vbExclamation
            Exit Sub
        End If

        原本.Copy After:=Sheets(Sheets.Count)
        Set 新シート = ActiveSheet

        シート名 = "チェックシート_" & 入力日

        If シート存在チェック(シート名) Then
            MsgBox "同じ日付のシートが既にあるためスキップします：" & vbCrLf & シート名, vbInformation

            Application.DisplayAlerts = False
            新シート.Delete
            Application.DisplayAlerts = True

            GoTo 次のループ
        End If

        新シート.Name = シート名

        On Error Resume Next
        新シート.Shapes("シート作成ボタン").Delete
        On Error GoTo 0

次のループ:
    Next i

End Sub

シート存在チェック関数:
Function シート存在チェック(ByVal 名前 As String) As Boolean
    On Error Resume Next
    シート存在チェック = Not Sheets(名前) Is Nothing
    On Error GoTo 0
End Function

```
---

## 📌 ポイント（Key Points）
- 重複日付のシートを作らないため、ミス防止に効果的
- 原本シートの UI を汚さないよう、不要なボタンは削除
- 実務でよくある「日付別フォーマット作成」を自動化
- シンプルな UI とコードで保守性も高い 

---

## 📂 想定利用シーン（Use Cases）
- 日次チェック作業のフォーマット作成
- 隔週、または月末・月初の大量シート生成
- チームでの作業フォーマット統一
- 手作業ミスを減らしたい現場向け

---

## 🌈 今後の拡張案（Future Enhancements）
- 日付入力をカレンダー選択に変更
- 自動で今日の日付を提案
- チェック処理（照合・判定）との連携
- UI ボタンの追加（必要に応じて）
