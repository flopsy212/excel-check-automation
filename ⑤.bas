Attribute VB_Name = "⑤"
Sub finalLastcheckfiles_macrosheet() ' 判定者シートの役割とEmail列だけ入力してAHEADと合致するか確認
    Dim Sheet1 As Worksheet
    Dim Sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim rowIdx As Long
    Dim RoleColumn2 As Long
    Dim EmailColumn2 As Long
    Dim mismatchDetails As String

    ' シート1を設定
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("判定者")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "やるやらシートがありません。", vbExclamation
        Exit Sub
    End If

    ' シート2を選択
    Dim selectedSheetIndex As Integer
    Dim SheetNameList As String
    SheetNameList = "シート名リスト:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        SheetNameList = SheetNameList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    selectedSheetIndex = CInt(InputBox("比較したいシートを選択してください。" & vbCrLf & SheetNameList))
    If selectedSheetIndex < 1 Or selectedSheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "無効な番号です。"
        Exit Sub
    End If

    Set Sheet2 = ThisWorkbook.Sheets(selectedSheetIndex)

    ' シート1とシート2の最終行を取得
    lastRow1 = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row

    ' シート2の「役割」と「Email」列を特定
    RoleColumn2 = 0
    EmailColumn2 = 0
    For ColIdx = 1 To Sheet2.Cells(1, Columns.Count).End(xlToLeft).Column
        If Trim(Sheet2.Cells(1, ColIdx).Value) = "役割" Then
            RoleColumn2 = ColIdx
        ElseIf Trim(Sheet2.Cells(1, ColIdx).Value) = "Email" Then
            EmailColumn2 = ColIdx
        End If
    Next ColIdx

    ' エラー処理
    If RoleColumn2 = 0 Or EmailColumn2 = 0 Then
        MsgBox "シート2に必要な列（役割またはEmail）が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 比較開始
    mismatchDetails = ""
    For rowIdx = 2 To Application.WorksheetFunction.Min(lastRow1, lastRow2)
        Dim Role1 As String, Email1 As String
        Dim Role2 As String, Email2 As String

        ' シート1とシート2の値を取得
        Role1 = Trim(Sheet1.Cells(rowIdx, 1).Value) ' シート1の役割 (A列)
        Email1 = Trim(Sheet1.Cells(rowIdx, 2).Value) ' シート1のEmail (B列)
        Role2 = Trim(Sheet2.Cells(rowIdx, RoleColumn2).Value) ' シート2の役割
        Email2 = Trim(Sheet2.Cells(rowIdx, EmailColumn2).Value) ' シート2のEmail

        ' 比較して不一致を記録
        If Role1 <> Role2 Or Email1 <> Email2 Then
            mismatchDetails = mismatchDetails & "行 " & rowIdx & " に不一致があります:" & vbCrLf & _
                             "  シート1 - 役割: " & Role1 & ", Email: " & Email1 & vbCrLf & _
                             "  シート2 - 役割: " & Role2 & ", Email: " & Email2 & vbCrLf & vbCrLf
        End If
    Next rowIdx

    ' 結果を新しいシートに出力
    If mismatchDetails <> "" Then
        Call WriteMismatchToNewSheet(mismatchDetails)
        MsgBox "不一致が見つかりました。詳細は新しいシートを確認してください。", vbInformation
    Else
        MsgBox "全て一致しています。", vbInformation
    End If
End Sub

Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' 新しいシートを追加
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行（ファイナル最終）"

    ' ヘッダーを書き込む
    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    ' MismatchRows を改行で分割して配列に格納
    Lines = Split(MismatchRows, vbCrLf)

    ' 不一致情報を書き込む
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            ' 元の文字列をそのまま書き込む
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub


