Sub Lastcheckfiles_macrosheet()
    '============================
    ' 判定者シートの「役割」と「連絡先」列を、
    ' 要件一覧ビューからエクスポートされた別シートと照合し、
    ' 行ごとの一致・不一致を確認するマクロ。
    '============================

    Dim Sheet1 As Worksheet ' 判定者シート（固定）
    Dim Sheet2 As Worksheet ' 比較対象のシート（ユーザー選択）
    Dim lastRow1 As Long, lastRow2 As Long
    Dim rowIdx As Long
    Dim RoleColumn2 As Long, 連絡先Column2 As Long
    Dim mismatchDetails As String

    '----------------------------------------
    ' Step1: 「判定者」シートを取得
    '----------------------------------------
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("判定者")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "統合シートがありません。", vbExclamation
        Exit Sub
    End If

    '----------------------------------------
    ' Step2: 比較対象シートを選択させる
    '----------------------------------------
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

    '----------------------------------------
    ' Step3: 各シートの最終行取得
    '----------------------------------------
    lastRow1 = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row

    '----------------------------------------
    ' Step4: シート2の「役割」と「連絡先」列の列番号を特定
    '----------------------------------------
    RoleColumn2 = 0
    連絡先Column2 = 0
    For ColIdx = 1 To Sheet2.Cells(1, Columns.Count).End(xlToLeft).Column
        If Trim(Sheet2.Cells(1, ColIdx).Value) = "役割" Then
            RoleColumn2 = ColIdx
        ElseIf Trim(Sheet2.Cells(1, ColIdx).Value) = "連絡先" Then
            連絡先Column2 = ColIdx
        End If
    Next ColIdx

    If RoleColumn2 = 0 Or 連絡先Column2 = 0 Then
        MsgBox "シート2に必要な列（役割または連絡先）が見つかりません。", vbExclamation
        Exit Sub
    End If

    '----------------------------------------
    ' Step5: 各行ごとに「役割」「連絡先」を比較
    '----------------------------------------
    mismatchDetails = ""
    For rowIdx = 2 To Application.WorksheetFunction.Min(lastRow1, lastRow2)
        Dim Role1 As String, 連絡先1 As String
        Dim Role2 As String, 連絡先2 As String

        ' 判定者シート（シート1）の値取得
        Role1 = Trim(Sheet1.Cells(rowIdx, 1).Value)   ' A列: 役割
        連絡先1 = Trim(Sheet1.Cells(rowIdx, 2).Value)  ' B列: 連絡先

        ' 比較対象シート（シート2）の値取得
        Role2 = Trim(Sheet2.Cells(rowIdx, RoleColumn2).Value)
        連絡先2 = Trim(Sheet2.Cells(rowIdx, 連絡先Column2).Value)

        ' 値が一致していなければ記録
        If Role1 <> Role2 Or 連絡先1 <> 連絡先2 Then
            mismatchDetails = mismatchDetails & "行 " & rowIdx & " に不一致があります:" & vbCrLf & _
                             "  シート1 - 役割: " & Role1 & ", 連絡先: " & 連絡先1 & vbCrLf & _
                             "  シート2 - 役割: " & Role2 & ", 連絡先: " & 連絡先2 & vbCrLf & vbCrLf
        End If
    Next rowIdx

    '----------------------------------------
    ' Step6: 結果を表示・出力
    '----------------------------------------
    If mismatchDetails <> "" Then
        Call WriteMismatchToNewSheet(mismatchDetails)
        MsgBox "不一致が見つかりました。詳細は新しいシートを確認してください。", vbInformation
    Else
        MsgBox "全て一致しています。", vbInformation
    End If
End Sub

'===========================
' 不一致行を書き出す共通サブ処理
'===========================
Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' 新しいシートを作成
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行（最終チェック）"

    ' ヘッダー行の出力
    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    ' 改行区切りで1行ずつ書き出し
    Lines = Split(MismatchRows, vbCrLf)
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub




