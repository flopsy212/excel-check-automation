Sub checkMacro()
'============================
' 「統合表」の「採否マーク列」と、

' 「要件一覧ビュー要件一覧表ビュー」の「仕向列」を比較して、
' 値が一致しない箇所を赤に染め、不一致内容を別シートに記録するマクロ
'============================
    Dim range1 As Range, range2 As Range ' 比較する2つの列範囲
    Dim cell1 As Range, cell2 As Range ' 比較対象のセル
    Dim mismatchDetails As String ' 不一致の詳細情報を記録する文字列
    Dim val1 As Variant, val2 As Variant ' 各セルの値
    Dim row1 As Long, row2 As Long ' 行番号（シート1とシート2）
    Dim startRow1 As Long, endRow1 As Long ' シート1で比較する範囲の開始・終了行
    Dim keyValue As String ' シート1を絞り込むためのキー値
    Dim Sheet1 As Worksheet, Sheet2 As Worksheet ' 比較対象の2シート
    Dim sheetList As String ' シート名の一覧（選択用）
    Dim sheetIndex1 As Integer, sheetIndex2 As Integer ' ユーザーが選択したシート番号
    Dim 管理IDCol As Long ' 「管理ID.」の列番号
    Dim 管理IDValue As String ' 「管理ID.」の値（先頭4桁）


'-------------------------------
' Step1: シート一覧を作ってユーザーに2枚選ばせる
'-------------------------------

    sheetList = "シート名リスト:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    ' ユーザーに比較するシート1を選択させる
    sheetIndex1 = Application.InputBox("統合シートを選択してください（番号を入力）:" & vbCrLf & sheetList, "シート選択", Type:=1)
    If sheetIndex1 < 1 Or sheetIndex1 > ThisWorkbook.Sheets.Count Then
        MsgBox "正しいシート番号を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' ユーザーに比較するシート2を選択させる
    sheetIndex2 = Application.InputBox("比較対象のシートを選択してください（番号を入力）:" & vbCrLf & sheetList, "シート選択", Type:=1)
    If sheetIndex2 < 1 Or sheetIndex2 > ThisWorkbook.Sheets.Count Then
        MsgBox "正しいシート番号を入力してください。", vbExclamation
        Exit Sub
    End If

    ' シートを設定
    Set Sheet1 = ThisWorkbook.Sheets(sheetIndex1)
    Set Sheet2 = ThisWorkbook.Sheets(sheetIndex2)
    
    

'-------------------------------
' Step2: 「管理ID.」列からキー値（先頭4桁）を取得
'-------------------------------
    
    On Error Resume Next
    管理IDCol = Application.Match("管理ID.", Sheet2.Rows(1), 0)
    On Error GoTo 0
    
    ' エラー処理
    If 管理IDCol = 0 Then
        MsgBox """管理ID.""" & " ラベルが見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' 「管理ID.」列の2行目の値の左から4文字を取得
    管理IDValue = Left(Sheet2.Cells(2, 管理IDCol).Value, 4)
    keyValue = Trim(管理IDValue)
    

'-------------------------------
' Step3: シート1のA列から、keyValueに一致する範囲を取得
'-------------------------------

    Dim foundCell As Range
    Set foundCell = Sheet1.Columns(1).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        MsgBox "キー値（" & keyValue & "）が見つかりませんでした。", vbExclamation
        Exit Sub
    Else
        startRow1 = foundCell.Row
    End If

    ' シート1のA列でキー値が終わる行を検索
    Set foundCell = Sheet1.Columns(1).Find(What:=keyValue & "*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        MsgBox "キー値（" & keyValue & "）の範囲が見つかりませんでした。", vbExclamation
        Exit Sub
    Else
        endRow1 = foundCell.Row
    End If
    
    
'-------------------------------
' Step4: 比較する列をユーザーに選ばせる
'-------------------------------
    
    ' シート1で比較する列を選択
    Set range1 = Application.InputBox("比較する列を選択してください（例: =" & Sheet1.Name & "!$K:$K）。", Type:=8)
    If range1 Is Nothing Or range1.Worksheet.Name <> Sheet1.Name Then
        MsgBox "比較元シートの列が正しく選択されていません。処理を終了します。", vbExclamation
        Exit Sub
    End If

    ' シート2で比較する列を選択
    Set range2 = Application.InputBox("比較する列を選択してください（例: =" & Sheet2.Name & "!$AN:$AN）。", Type:=8)
    If range2 Is Nothing Then
        MsgBox "比較対象シートの列が選択されませんでした。処理を終了します。", vbExclamation
        Exit Sub
    End If
    
    
'-------------------------------
' Step5: 比較処理（1対1で行を比較）
'-------------------------------

    ' 初期化
    mismatchDetails = ""

    ' 比較処理（1対1の行比較）
    row2 = 2 ' シート2の開始行
    For row1 = startRow1 To endRow1
        If row2 > Sheet2.Cells(Sheet2.Rows.Count, range2.Column).End(xlUp).Row Then Exit For

        ' セルを取得
        Set cell1 = Sheet1.Cells(row1, range1.Column)
        Set cell2 = Sheet2.Cells(row2, range2.Column)

        ' 値を取得
        val1 = Trim(CStr(cell1.Value))
        val2 = Trim(CStr(cell2.Value))

        ' 背景色と値の比較
        If cell1.Interior.Color = RGB(169, 169, 169) Or cell2.Interior.Color = RGB(169, 169, 169) Or _
        cell1.Interior.Color = RGB(166, 166, 166) Or cell2.Interior.Color = RGB(166, 166, 166) Or _
        cell1.Interior.Color = RGB(0, 0, 0) Or cell2.Interior.Color = RGB(0, 0, 0) Then
        
            GoTo ContinueLoop
        Else
        
        ' 両方空白ならエラー扱い
        If val1 = "" And val2 = "" Then
            cell1.Interior.Color = RGB(255, 0, 0)
            cell2.Interior.Color = RGB(255, 0, 0)
            mismatchDetails = mismatchDetails & "シート1行 " & row1 & " / シート2行 " & row2 & ": 両方空白です。" & vbCrLf
            GoTo ContinueLoop
        End If
        
        ' 白塗りセル：完全一致を確認
        If val1 <> val2 Then
            cell1.Interior.Color = RGB(255, 0, 0)
            cell2.Interior.Color = RGB(255, 0, 0)
            mismatchDetails = mismatchDetails & "シート1行 " & row1 & " / シート2行 " & row2 & ": 白塗りセルで不一致 (Cell1: [" & val1 & "], Cell2: [" & val2 & "])" & vbCrLf
        End If
    End If
    
ContinueLoop:
    
        ' シート2の次の行へ
        row2 = row2 + 1
    Next row1
    
    
'-------------------------------
' Step6: 結果を表示 or 新シートに出力
'-------------------------------

    If mismatchDetails = "" Then
        MsgBox "すべて一致しました！", vbInformation
    Else
        MsgBox "以下の不一致が見つかりました:" & vbCrLf & mismatchDetails, vbExclamation
        ' 不一致行を新しいシートに書き出す
        WriteMismatchResults mismatchDetails
    End If
End Sub


'============================
' 不一致内容を新しいシートに書き出す処理
'============================

Sub WriteMismatchResults(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' 新しいシートを追加
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行(チェックマクロ)"

    ' ヘッダーを書き込む
    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    ' MismatchRows を改行で分割して配列に格納
    Lines = Split(MismatchRows, vbCrLf)

    ' 不一致情報を書き込む
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub


