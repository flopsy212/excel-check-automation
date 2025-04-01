Sub compareTasksColumns()
    '============================
    ' 任意列の最終チェック
    ' 統合シートと他の任意のシートの特定列を比較し、
    ' 不一致があればセルを赤く塗り、不一致内容を別シートに記録する。
    '============================
    Dim range1 As Range, range2 As Range
    Dim cell1 As Range, cell2 As Range
    Dim mismatchDetails As String
    Dim val1 As Variant, val2 As Variant
    Dim row1 As Long, row2 As Long
    Dim startRow1 As Long, endRow1 As Long
    Dim keyValue As String
    Dim Sheet1 As Worksheet, Sheet2 As Worksheet
    Dim P2Value As String
    Dim sheetList As String, sheetIndex As Integer
    Dim lastRow2 As Long
    Dim 管理IDCol As Long
    Dim 管理IDValue As String

    '----------------------------------------
    ' Step1: 「統合」シートの取得
    '----------------------------------------
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("統合")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "「統合」シートが見つかりません。処理を終了します。", vbExclamation
        Exit Sub
    End If

    '----------------------------------------
    ' Step2: 比較対象のシートをユーザーに選ばせる
    '----------------------------------------
    sheetList = "シート名リスト:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    sheetIndex = Application.InputBox("比較するシートを選択してください（番号を入力）:" & vbCrLf & sheetList, "シート選択", Type:=1)
    If sheetIndex < 1 Or sheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "正しいシート番号を入力してください。", vbExclamation
        Exit Sub
    End If
    Set Sheet2 = ThisWorkbook.Sheets(sheetIndex)

    '----------------------------------------
    ' Step3: 比較対象シートから「管理ID.」列を探し、キー値を取得
    '----------------------------------------
    管理IDCol = Application.Match("管理ID.", Sheet2.Rows(1), 0)
    If Not IsError(管理IDCol) Then
        管理IDValue = Left(Sheet2.Cells(2, 管理IDCol).Value, 4)
    Else
        MsgBox """管理ID.""" & " ラベルが見つかりませんでした。"
        Exit Sub
    End If

    keyValue = Trim(管理IDValue)

    '----------------------------------------
    ' Step4: 統合シートのA列から、キー値で行範囲を特定
    '----------------------------------------
    startRow1 = Sheet1.Columns(1).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole).Row
    endRow1 = Sheet1.Columns(1).Find(What:=keyValue & "*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row

    If startRow1 = 0 Or endRow1 = 0 Then
        MsgBox "キー値（" & keyValue & "）が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    '----------------------------------------
    ' Step5: 比較したい列（統合・比較シート）をユーザーに選ばせる
    '----------------------------------------
    Set range1 = Application.InputBox("比較する列を統合シートで選択してください（例: =統合!$B:$B）。", Type:=8)
    If range1 Is Nothing Or range1.Worksheet.Name <> "統合" Then
        MsgBox "統合シートの列が正しく選択されていません。処理を終了します。", vbExclamation
        Exit Sub
    End If

    Set range2 = Application.InputBox("比較する列を他のシートで選択してください（例: =" & Sheet2.Name & "!$P:$P）。", Type:=8)
    If range2 Is Nothing Then
        MsgBox "シート2の列が選択されませんでした。処理を終了します。", vbExclamation
        Exit Sub
    End If

    '----------------------------------------
    ' Step6: 比較処理（1対1で行ごとに比較）
    '----------------------------------------
    mismatchDetails = ""
    row2 = 2 ' 比較シートの2行目から

    For row1 = startRow1 To endRow1
        If row2 > Sheet2.Cells(Sheet2.Rows.Count, range2.Column).End(xlUp).Row Then Exit For

        Set cell1 = Sheet1.Cells(row1, range1.Column)
        Set cell2 = Sheet2.Cells(row2, range2.Column)

        val1 = Trim(CStr(cell1.Value))
        val2 = Trim(CStr(cell2.Value))

        If val1 <> val2 Then
            cell1.Interior.Color = RGB(255, 0, 0)
            cell2.Interior.Color = RGB(255, 0, 0)
            mismatchDetails = mismatchDetails & "シート1行 " & row1 & " / シート2行 " & row2 & ": 値が一致しません (Cell1: [" & val1 & "], Cell2: [" & val2 & "])" & vbCrLf
        End If

        row2 = row2 + 1
    Next row1

    '----------------------------------------
    ' Step7: 比較結果の出力
    '----------------------------------------
    If mismatchDetails = "" Then
        MsgBox "すべて一致しました！", vbInformation
    Else
        MsgBox "以下の不一致が見つかりました:" & vbCrLf & mismatchDetails, vbExclamation
        Call WriteMismatchToNewSheet(mismatchDetails)
    End If
End Sub

'===========================
' 不一致行を新しいシートに書き出すサブ処理
'===========================
Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行(任意列の最終チェック)"

    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    Lines = Split(MismatchRows, vbCrLf)
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub

