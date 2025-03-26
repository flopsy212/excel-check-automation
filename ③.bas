Attribute VB_Name = "③"
Sub compareyaruyara_hanteiyouhi() '室課,判定要否列の確認
    Dim requirementNumber As String ' 要件番号を格納するための文字列変数
    Dim requirementRow As Long ' 要件番号が見つかった行番号を格納するための変数
    Dim roomKaa As String ' 室課（おそらく部門や担当）の情報を格納するための文字列変数
    Dim yColumnValue As String ' Y列の値を格納するための文字列変数
    Dim roomColumn As Long ' 室課列の列番号を格納するための変数
    Dim judgementColumn As Long ' 判定列の列番号を格納するための変数
    Dim lastRow1 As Long ' 最初のシート（Sheet1）におけるデータの最終行を格納する変数
    Dim lastRow2 As Long ' 2番目のシート（Sheet2）におけるデータの最終行を格納する変数
    Dim matchResult As Variant ' Match関数の結果（検索された位置）を格納するための変数
    Dim rowNumber As Long ' 行番号を格納するための変数
    Dim Sheet1 As Worksheet ' 最初のシート（Sheet1）を格納するためのWorksheetオブジェクト
    Dim Sheet2 As Worksheet ' 2番目のシート（Sheet2）を格納するためのWorksheetオブジェクト
    Dim MismatchRows As String ' 不一致行を格納するための文字列変数
    Dim selectedSheetIndex As Long

    MismatchRows = "" ' 初期化

    ' シート1を「やるやら」という名前で固定
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("やるやら")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "やるやらシートがありません。", vbExclamation
        Exit Sub
    End If

    ' ユーザーに要件番号を選択させる
    lastRow1 = Sheet1.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row
    Dim requirementList As String
    Dim i As Long
    For i = 2 To lastRow1
        requirementList = requirementList & Sheet1.Cells(i, "A").Value & vbCrLf
    Next i

    requirementNumber = InputBox("以下の要件番号から選択してください:" & vbCrLf & requirementList, "要件番号選択")

    ' 要件番号が入力されていない場合は終了
    If Len(Trim(requirementNumber)) = 0 Then
        MsgBox "要件番号が入力されていません。", vbExclamation
        Exit Sub
    End If

    ' シート1のA列を検索して一致する要件番号を探す
    matchResult = Application.Match(requirementNumber, Sheet1.Range("A2:A" & lastRow1), 0)

    ' 一致しない場合、エラーメッセージを表示して終了
    If IsError(matchResult) Then
        MsgBox "指定された要件番号はシート1のA列に存在しません。", vbExclamation
        Exit Sub
    End If

    ' 一致した場合、その行の番号を取得
    requirementRow = matchResult + 1 ' matchResultは1から始まるため調整

    ' シート名リストの作成
    Dim sheetList As String
    sheetList = "シート名リスト:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    ' シート2を選択（番号で選ぶ）
    selectedSheetIndex = CInt(InputBox("比較したいシートを選択してください:" & vbCrLf & sheetList))
    If selectedSheetIndex < 1 Or selectedSheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "無効な番号です。"
        Exit Sub
    End If
    Set Sheet2 = ThisWorkbook.Sheets(selectedSheetIndex)

    ' シート1の1行目から「室課」と「判定要否」の列を検索
    roomColumn = Application.Match("室課", Sheet1.Rows(1), 0)
    judgementColumn = Application.Match("判定要否", Sheet1.Rows(1), 0)

    ' 「室課」列と「判定要否」列が見つからない場合、エラーを出力して終了
    If IsError(roomColumn) Then
        MsgBox """室課""列が見つかりませんでした。"
        Exit Sub
    End If
    If IsError(judgementColumn) Then
        MsgBox """判定要否""列が見つかりませんでした。"
        Exit Sub
    End If

    ' シート2の最終行を取得
    lastRow2 = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row

    ' シート2の2行目から最終行までループ
    For rowNumber = 2 To lastRow2
        ' シート2のA列（室課）を取得
        roomKaa = Sheet2.Cells(rowNumber, 1).Value

        ' シート1の一致する要件番号の行で「室課」の値を取得
        If Sheet1.Cells(requirementRow, roomColumn).Value = roomKaa Then
            ' 室課が一致した場合、判定要否列の値を取得
            yColumnValue = Sheet1.Cells(requirementRow, judgementColumn).Value

            ' シート2のC列と一致しているか確認
            If Trim(yColumnValue) = "" And Trim(Sheet2.Cells(rowNumber, 3).Value) = "" Then
                ' 両方が空白の場合は何もしない
            ElseIf yColumnValue = Sheet2.Cells(rowNumber, 3).Value Then
                ' 一致する場合は何もしない
            Else
                ' 不一致の場合、シート1の判定要否列とシート2のC列を赤色で塗りつぶす
                Sheet1.Cells(requirementRow, judgementColumn).Interior.Color = RGB(255, 0, 0) ' シート1の判定要否列を赤色
                Sheet2.Cells(rowNumber, 3).Interior.Color = RGB(255, 0, 0) ' シート2のC列を赤色
                ' 不一致行の情報を収集
                MismatchRows = MismatchRows & "シート1行" & requirementRow & "とシート2行" & rowNumber & vbCrLf
            End If
        End If
    Next rowNumber

    ' 不一致行が存在する場合、新しいシートにメモを作成
    If MismatchRows <> "" Then
        Call WriteMismatchToNewSheet(MismatchRows)
    Else
        MsgBox "不一致は見つかりませんでした。"
    End If
End Sub

Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' 新しいシートを追加
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行（判定要否）"

    ' ヘッダーを書き込む
    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    ' mismatchRows を改行で分割して配列に格納
    Lines = Split(MismatchRows, vbCrLf)

    ' 不一致情報を書き込む
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex

    ' 完了通知
    MsgBox "不一致行の詳細が新しいシートに保存されました。"
End Sub


   
