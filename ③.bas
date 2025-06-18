
Sub compareTasks_decisionFlag()
    '============================
    ' 要件一覧ビューの判定要否の確認項目からファイルをエクスポートしてきたシートと
    '「統合」シートとの「部署」「判定要否」の一致を確認するマクロ。
    ' 不一致があれば、セルを赤色にして、別シートに記録します。
    '============================
    
    Dim requirementRow As Long ' 要件番号がある行番号（現在は未使用）
    Dim roomKaa As String ' 比較シートの部署の値
    Dim yColumnValue As String ' 統合シートの判定要否の値
    Dim roomColumn As Long ' 部署の列番号（統合）
    Dim judgementColumn As Long ' 判定要否の列番号（統合）
    Dim lastRow1 As Long ' 統合シートの最終行
    Dim lastRow2 As Long ' 比較シートの最終行
    Dim matchResult As Variant ' Match関数の結果（見つかった行位置）（現在は未使用）
    Dim rowNumber As Long ' ループ用の行番号
    Dim Sheet1 As Worksheet ' 「統合」シート
    Dim Sheet2 As Worksheet ' 比較対象のシート
    Dim MismatchRows As String ' 不一致行の記録
    Dim selectedSheetIndex As Long ' ユーザーが選んだシート番号
    Dim i As Long

    MismatchRows = "" ' 不一致行の記録を初期化

    '------------------------------------
    ' Step1: 「統合」シートがあるかチェック
    '------------------------------------
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("統合")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "統合シートがありません。", vbExclamation
        Exit Sub
    End If
    lastRow1 = Sheet1.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row

    '------------------------------------
    ' Step2: 比較するシートを選ばせる
    '------------------------------------
    Dim sheetList As String
    sheetList = "シート名リスト:" & vbCrLf
    For i = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & i & ". " & ThisWorkbook.Sheets(i).Name & vbCrLf
    Next i

    selectedSheetIndex = CInt(InputBox("比較したいシートを選択してください:" & vbCrLf & sheetList))
    If selectedSheetIndex < 1 Or selectedSheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "無効な番号です。"
        Exit Sub
    End If
    Set Sheet2 = ThisWorkbook.Sheets(selectedSheetIndex)

    '------------------------------------
    ' Step3: 「部署」と「判定要否」列の番号を特定
    '------------------------------------
    roomColumn = Application.Match("部署", Sheet1.Rows(1), 0)
    judgementColumn = Application.Match("判定要否", Sheet1.Rows(1), 0)

    If IsError(roomColumn) Then
        MsgBox """部署""列が見つかりませんでした。"
        Exit Sub
    End If
    If IsError(judgementColumn) Then
        MsgBox """判定要否""列が見つかりませんでした。"
        Exit Sub
    End If

    '------------------------------------
    ' Step4: 比較シートの全行と比較処理開始
    '------------------------------------
    lastRow2 = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow1
        For rowNumber = 2 To lastRow2
            roomKaa = Sheet2.Cells(rowNumber, 1).Value

            If Sheet1.Cells(i, roomColumn).Value = roomKaa Then
                yColumnValue = Sheet1.Cells(i, judgementColumn).Value

                If Trim(yColumnValue) = "" And Trim(Sheet2.Cells(rowNumber, 3).Value) = "" Then
                    ' 両方空欄 → OK
                ElseIf yColumnValue = Sheet2.Cells(rowNumber, 3).Value Then
                    ' 一致 → OK
                Else
                    ' 不一致 → 赤く塗る
                    Sheet1.Cells(i, judgementColumn).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(rowNumber, 3).Interior.Color = RGB(255, 0, 0)
                    MismatchRows = MismatchRows & "統合行" & i & " と " & Sheet2.Name & " 行" & rowNumber & vbCrLf
                End If
            End If
        Next rowNumber
    Next i

    '------------------------------------
    ' Step5: 不一致があれば、新しいシートに記録
    '------------------------------------
    If MismatchRows <> "" Then
        Call WriteMismatchToNewSheet(MismatchRows)
    Else
        MsgBox "不一致は見つかりませんでした。"
    End If
End Sub

'===========================
' 不一致行を新しいシートに書き出すサブ処理
'===========================
Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' 新しいシートを作成
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行（判定要否）"

    ' ヘッダー行
    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    ' 改行ごとに分割して書き出し
    Lines = Split(MismatchRows, vbCrLf)
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex

    MsgBox "不一致行の詳細が新しいシートに保存されました。"
End Sub


