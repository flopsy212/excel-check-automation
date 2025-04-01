Sub ProcessSheets() 'シート改変
    Dim ws As Worksheet '改変するシートを格納する変数
    Dim yaruyaraSheet As Worksheet '統合するシートをyaruyarasheetとして定義
    Dim headerSourceSheet As Worksheet
    Dim columnsToKeep As Variant

    ' やるやらシートを作成
    On Error Resume Next
    Set yaruyaraSheet = ThisWorkbook.Sheets("やるやら")
    On Error GoTo 0

    If yaruyaraSheet Is Nothing Then
        Set yaruyaraSheet = ThisWorkbook.Sheets.Add
        yaruyaraSheet.Name = "やるやら"
    End If

    ' やるやらシートと改変するシートの設定
    Set yaruyaraSheet = ThisWorkbook.Sheets("やるやら")
    Set headerSourceSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' 必要な列のリスト（D, M, N, P, AM, AV, AY）'必要な列リストを今回は固定で行っているが、3SLの場合以外は列の場所が違うとエラーが出る
    columnsToKeep = Array(4, 13, 14, 16, 39, 48, 51)

    ' シート名変更処理を呼び出し
    Call RenameSheetsBasedOnA0No

    ' Sheet1などのシート以外をループして処理
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Sheet1" And ws.Name <> "全体フロー" And ws.Name <> "手順説明" And ws.Name <> "判定者" And ws.Name <> "やるやら" And ws.Name <> "Innovator" And ws.Name <> "見本" And ws.Name <> "Innovator (2)" Then
            ' 処理フロー
            Call ProcessSheetColumns(ws, columnsToKeep)
            Call ProcessAdditionalColumns(ws)
            Call FormatAndStyleSheet(ws)
            Call InsertFormulasAndFormatting(ws, yaruyaraSheet)
            Call CopyDataToYaruyara(ws, yaruyaraSheet)
            Call RestrictInputBasedOnColumns(yaruyaraSheet)
        End If
    Next ws

    ' やるやらシートの仕上げ
    Call SetupYaruyaraSheet(yaruyaraSheet, headerSourceSheet)
End Sub

' 以下略（他のSubも必要なら続き貼ってください）
