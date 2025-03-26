Attribute VB_Name = "①"
Sub ProcessSheets() 'シート改変
    Dim ws As Worksheet '改変するシートを格納する変数
    Dim yaruyaraSheet As Worksheet '統合するシートをyaruyarasheetとして定義
    Dim headerSourceSheet As Worksheet
    Dim columnsToKeep As Variant
    Dim col As Variant

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
     
     ' ラベル名リスト
    labelNames = Array("Title EN", "分類名", "A要件名1", "A0 No.", "採否マーク1", "室課", "判定ランク")
    Set columnsToKeep = New Collection
    
    For Each col In columnsToKeep
    Debug.Print col
    Next col

    ' ヘッダー行（例えば1行目）を検索して、対応する列番号を取得
    Set headerRange = headerSourceSheet.Rows(1)

    ' ラベル名に基づいて列番号を検索し、Collectionに格納
    For Each Label In labelNames
        Set cell = headerRange.Find(Label, LookIn:=xlValues, LookAt:=xlWhole)
        If Not cell Is Nothing Then
            colNum = cell.Column
            columnsToKeep.Add colNum ' 見つかった列番号を追加
        Else
            MsgBox "ラベル " & Label & " が見つかりません。"
        End If
    Next Label

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

Sub RenameSheetsBasedOnA0No()
    Dim ws As Worksheet
    Dim cellValue As Variant
    Dim newName As String
    Dim colA0No As Variant
            
    ' シートをループしてシート名を変更
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Sheet1" And ws.Name <> "全体フロー" And ws.Name <> "手順説明" And ws.Name <> "判定者" And ws.Name <> "やるやら" And ws.Name <> "Innovator" And ws.Name <> "見本" And ws.Name <> "Innovator (2)" Then

            ' A0 No.列を検索
            colA0No = Application.Match("A0 No.", ws.Rows(1), 0)
            
            ' A0 No.が見つかった場合
            If Not IsError(colA0No) Then
                ' A0 No.列の2行目の値を確認
                cellValue = ws.Cells(2, colA0No).Value
            Else
                MsgBox """A0 No.""" & " ラベルが見つかりませんでした。"
            End If

            ' シート名を変更（エラーが発生しないように処理）
            On Error Resume Next
            ws.Name = newName
            On Error GoTo 0
        End If
    Next ws
End Sub

Sub ProcessSheetColumns(ws As Worksheet, columnsToKeep As Variant) '不要列削除・要件番号列挿入
    Dim col As Long
    Dim colReqNo As Variant
    Dim colA0No As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim colItem As Variant
    Dim found As Boolean
    
    For col = ws.UsedRange.Columns.Count To 1 Step -1
        found = False
        ' columnsToKeep のすべての値と比較
        For Each colItem In columnsToKeep
            If ws.Cells(1, col).Value = colItem Then
                found = True
                Exit For
            End If
        Next colItem
    
        ' 見つからなかった場合、列を削除
        If Not found Then ws.Columns(col).Delete
    Next col
    
    ' A列を挿入
    ws.Columns("A").Insert Shift:=xlToRight
    ws.Columns("A").NumberFormat = "@"

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' シート名を文字列として入力
    For i = 2 To lastRow
        ws.Cells(i, 1).Value = "'" & ws.Name
    Next i

    ' ヘッダーを設定
    ws.Cells(1, 1).Value = "要件番号"

    ' エラーチェックを無効化（必要に応じて）
    Application.ErrorCheckingOptions.NumberAsText = False

    ' 要件番号列を検索
    colReqNo = Application.Match("要件番号", ws.Rows(1), 0)

    If Not IsError(colReqNo) Then
        ' A0 No.列を検索
        colA0No = Application.Match("A0 No.", ws.Rows(1), 0)
        
        If Not IsError(colA0No) Then
            ' A0 No.列を要件番号列の右隣にコピーして挿入
            ws.Columns(colA0No).Copy
            ws.Columns(colReqNo + 1).Insert Shift:=xlToRight
            
            ' コピー後、元のA0 No.列を削除
            ws.Columns(colA0No + 1).Delete Shift:=xlToLeft
        Else
            MsgBox """A0 No.""" & " ラベルが見つかりませんでした。"
        End If
    Else
        MsgBox """要件番号""" & " ラベルが見つかりませんでした。"
    End If
End Sub

Sub ProcessAdditionalColumns(ws As Worksheet) '列挿入・室課列移動
    Dim colAheadInput As Variant
    Dim colRoom As Variant

    ' 「AHEAD入力可否」列の番号を検索
    ws.Cells(1, 23).Value = "AHEAD入力可否"
    colAheadInput = Application.Match("AHEAD入力可否", ws.Rows(1), 0)

    If Not IsError(colAheadInput) Then
        ' 「室課」列を検索
        colRoom = Application.Match("室課", ws.Rows(1), 0)

        If Not IsError(colRoom) Then
            ' 室課列を右隣に移動（挿入）
            ws.Columns(colRoom).Cut
            ws.Columns(colAheadInput + 1).Insert Shift:=xlToRight
        Else
            MsgBox """室課""" & " ラベルが見つかりませんでした。"
        End If
    Else
        MsgBox """AHEAD入力可否""" & " ラベルが見つかりませんでした。"
    End If
    
       ' 採否マーク列の左側に室課5列挿入
    ws.Columns("F:J").Insert Shift:=xlToRight
    ws.Range("F2:J" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).Value = ""

    ' 採否判定理由列の左側に室課5列挿入
    ws.Columns("M:Q").Insert Shift:=xlToRight
    ws.Range("M2:Q" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).Value = ""

    ' 挿入した列をグループ化
    ws.Columns("F:J").Group
    ws.Columns("M:Q").Group
End Sub

Sub FormatAndStyleSheet(ws As Worksheet)
    Dim colorToApply As Long
    Dim whereESS As Long
    Dim lastRow As Long
    Dim cell As Range

    ' シートの1行目にラベルを格納
    Dim headers As Variant
    headers = Array("BAT性能", "QJB MJB", "構造", "ESS熱マネ", "BTS熱マネ", "", "BAT性能", "QJB MJB", "構造", "ESS熱マネ", "BTS熱マネ", "採否判定理由", "採否理由チェック (委託元)", "採否理由チェック (委託先)", "承認アイテムA要件付表", "判定ランク", "", "", "機種担当室課", "判定要否")
    
    Dim i As Integer
    For i = LBound(headers) To UBound(headers)
        If headers(i) <> "" Then
            ws.Cells(1, i + 6).Value = headers(i)
        End If
    Next i

    ' 1行目で色が入っているセルを検索し、その色を他のセルに適用
    For Each cell In ws.Rows(1).Cells
        If cell.Interior.ColorIndex <> -4142 Then
            colorToApply = cell.Interior.Color ' 色を取得
            Exit For
        End If
    Next cell

    ' 色を他の1行目のセルに適用
    For Each cell In ws.Rows(1).Cells
        If cell.Interior.ColorIndex = -4142 Then
            cell.Interior.Color = colorToApply
        End If
    Next cell

    ' 室課側で入力が完了していないセルに黄色を付ける
    whereESS = 23
    lastRow = ws.Cells(ws.Rows.Count, whereESS).End(xlUp).Row
    For Each cell In ws.Range(ws.Cells(2, whereESS), ws.Cells(lastRow, whereESS))
        If Trim(cell.Value) <> "" And InStr(UCase(Trim(cell.Value)), "ESS") = 0 Then
            cell.EntireRow.Interior.Color = RGB(169, 169, 169)
        End If
    Next cell

    ' 特定列の1行目を黄色で塗りつぶす
    ws.Range("F1:J1").Interior.Color = RGB(255, 255, 0)
    ws.Range("L1:Q1").Interior.Color = RGB(255, 255, 0)
    ws.Range("T1:U1").Interior.Color = RGB(255, 255, 0)

    ' フィルターを設定
    ws.Rows(1).AutoFilter

    ' 列幅を自動調整
    ws.Columns.AutoFit

    ' すべてのシートにA列からY列まで太い格子をつける
    With ws.Range("A1:Y" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With
End Sub

Sub InsertFormulasAndFormatting(ws As Worksheet, yaruyaraSheet As Worksheet)
    Application.Calculation = xlCalculationAutomatic
    Dim cell As Range
    Dim lastRow As Long
    Dim targetRange As Range

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "データがありません。", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' シート保護解除
    ws.Unprotect

    ' 再計算を強制
    Application.Calculate

    ' K列に数式を設定 (F列からJ列に変更があった場合でも再計算)
    For Each cell In ws.Range("K2:K" & lastRow)
        ' 常に数式を再設定
        cell.FormulaR1C1 = _
           "=IF(COUNTA(RC[-5]:RC[-1])=0,"""",IF(COUNTIF(RC[-5]:RC[-1],""〇"")>0,""〇"",IF(COUNTIF(RC[-5]:RC[-1],""×"")>0,""×"",""-"")))"
    Next cell

    ' Z列に数式を追加
    For Each cell In ws.Range("Z2:Z" & lastRow)
        If Len(Trim(cell.Value)) = 0 Or Not cell.HasFormula Then
            cell.Formula = "=IF(OR(K" & cell.Row & "=""-"", K" & cell.Row & "=""""), ""全て該当せず"", IF(K" & cell.Row & "=""×"", ""全てテスト・確認せず"", IF(K" & cell.Row & "=""〇"", ""テスト要"", """")))"
        End If
    Next cell

    ' Y列に数式を追加
    For Each cell In ws.Range("Y2:Y" & lastRow)
        If Len(Trim(cell.Value)) = 0 Or Not cell.HasFormula Then
            cell.Formula = "=IFERROR(IF(Z" & cell.Row & "=""テスト要"", ""テスト要"", IF(Z" & cell.Row & "=""全てテスト・確認せず"", IF(IFERROR(COUNTIFS(W$1:W$10001, W" & cell.Row & ", Z$1:Z$10001, ""テスト要""), 0)>0, ""テスト要"", ""全てテスト・確認せず""), IF(Z" & cell.Row & "=""全て該当せず"", IF(IFERROR(COUNTIFS(W$1:W$10001, W" & cell.Row & ", Z$1:Z$10001, ""テスト要""), 0)>0, ""テスト要"", IF(IFERROR(COUNTIFS(W$1:W$10001, W" & cell.Row & ", Z$1:Z$10001, ""全てテスト・確認せず""), 0)>0, ""全てテスト・確認せず"", ""全て該当せず"")), Z" & cell.Row & "))), """")"
        End If
    Next cell
    
    ' V列に条件付き書式を追加
    Set targetRange = ws.Range("A2:Z" & lastRow)
    With targetRange
        .FormatConditions.Delete ' 既存の条件付き書式を削除
        ' 数式を利用した条件付き書式を追加
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$V2=""〇""")
            .Interior.Color = RGB(255, 255, 255) ' 白色に設定
        End With
    End With
        
    ' Z列を非表示
    ws.Columns("Z").Hidden = True
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    On Error GoTo 0
End Sub

Sub CopyDataToYaruyara(ws As Worksheet, yaruyaraSheet As Worksheet) 'やるやらシートへのコピペ準備
    Dim yaruyaraLastRow As Long
    Dim lastRow As Long
    Dim dataRange As Range
    Dim destRange As Range
    
    ' データの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 2行目以降のデータ範囲を設定
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Columns.Count))

    ' やるやらシートの最終行を取得
    yaruyaraLastRow = yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row

    ' 貼り付け先の範囲を設定
    Set destRange = yaruyaraSheet.Cells(yaruyaraLastRow + 1, 1).Resize(dataRange.Rows.Count, dataRange.Columns.Count)

    ' データとフォーマットをコピー
    dataRange.Copy
    destRange.PasteSpecial Paste:=xlPasteAll

    ' コピー後にクリップボードをクリア
    Application.CutCopyMode = False
    
    ' やるやらシートでも中央揃えを適用
    destRange.HorizontalAlignment = xlCenter
End Sub

Sub SetupYaruyaraSheet(yaruyaraSheet As Worksheet, headerSourceSheet As Worksheet)
    Dim lastRow As Long
    Dim targetRange As Range
    
    ' 保護を解除
    If yaruyaraSheet.ProtectContents Then
        yaruyaraSheet.Unprotect Password:="password" ' パスワードは適宜設定
    End If

    ' ヘッダー1行目をやるやらシートにコピー
    headerSourceSheet.Rows(1).Copy Destination:=yaruyaraSheet.Rows(1)

    ' やるやらシートの1行目にフィルターを設定
    yaruyaraSheet.Rows(1).AutoFilter

    ' A列からY列まで太い格子をつける
    With yaruyaraSheet.Range("A1:Y" & yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With

    ' 挿入した列をグループ化
    yaruyaraSheet.Columns("F:J").Group
    yaruyaraSheet.Columns("L:P").Group

    ' 最終行を取得
    lastRow = yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row

    ' V列の条件付き書式を設定
    Set targetRange = yaruyaraSheet.Range("A2:Z" & lastRow)
    With targetRange
        .FormatConditions.Delete ' 既存の条件付き書式を削除
        ' 数式を利用した条件付き書式を追加
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$V2=""〇""")
            .Interior.Color = RGB(255, 255, 255) ' 白色に設定
        End With
    End With

    ' セルの内容を水平方向に中央揃え
    yaruyaraSheet.Range("A1:Y" & yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row).HorizontalAlignment = xlCenter

    ' 列幅を自動調整
    yaruyaraSheet.Columns.AutoFit

    ' Z列を非表示
    yaruyaraSheet.Columns("Z").Hidden = True

    ' A列を昇順にソート
    With yaruyaraSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=yaruyaraSheet.Range("A2:A" & lastRow), Order:=xlAscending
        .SetRange yaruyaraSheet.Range("A1:Y" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' 保護を再設定
    yaruyaraSheet.Protect Password:="password", UserInterfaceOnly:=True

    MsgBox "処理が完了しました！"
End Sub

Sub RestrictInputBasedOnColumns(ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim qFilled As Boolean
    Dim kFilled As Boolean

    ' 保護解除
    If ws.ProtectContents Then
        ws.Unprotect Password:="password"
    End If

    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 入力制限を設定
    For Each cell In ws.Range("A2:Z" & lastRow).Cells
        If Not ((cell.Column >= 6 And cell.Column <= 10) Or _
                (cell.Column >= 12 And cell.Column <= 17) Or _
                (cell.Column >= 20 And cell.Column <= 21)) Then
            cell.Locked = True
        Else
            cell.Locked = False
        End If
    Next cell

    ' Q列とK列が埋まった場合のチェック
    qFilled = WorksheetFunction.CountBlank(ws.Range("Q2:Q" & lastRow)) = 0
    kFilled = WorksheetFunction.CountBlank(ws.Range("K2:K" & lastRow)) = 0

    If qFilled And kFilled Then
        ' Q列とK列が埋まったらV列とX列を解放
        ws.Range("V2:V" & lastRow).Locked = False
        ws.Range("X2:X" & lastRow).Locked = False
    Else
        ' 再ロック
        ws.Range("V2:V" & lastRow).Locked = True
        ws.Range("X2:X" & lastRow).Locked = True
    End If

    ' シート保護を再設定
    ws.Protect Password:="password", UserInterfaceOnly:=True, AllowFiltering:=True
End Sub


