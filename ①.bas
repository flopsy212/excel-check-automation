Attribute VB_Name = "�@"
Sub ProcessSheets() '�V�[�g����
    Dim ws As Worksheet '���ς���V�[�g���i�[����ϐ�
    Dim yaruyaraSheet As Worksheet '��������V�[�g��yaruyarasheet�Ƃ��Ē�`
    Dim headerSourceSheet As Worksheet
    Dim columnsToKeep As Variant
    Dim col As Variant

    ' �����V�[�g���쐬
    On Error Resume Next
    Set yaruyaraSheet = ThisWorkbook.Sheets("�����")
    On Error GoTo 0

    If yaruyaraSheet Is Nothing Then
        Set yaruyaraSheet = ThisWorkbook.Sheets.Add
        yaruyaraSheet.Name = "�����"
    End If

    ' �����V�[�g�Ɖ��ς���V�[�g�̐ݒ�
    Set yaruyaraSheet = ThisWorkbook.Sheets("�����")
    Set headerSourceSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
     
     ' ���x�������X�g
    labelNames = Array("Title EN", "���ޖ�", "A�v����1", "A0 No.", "�̔ۃ}�[�N1", "����", "���胉���N")
    Set columnsToKeep = New Collection
    
    For Each col In columnsToKeep
    Debug.Print col
    Next col

    ' �w�b�_�[�s�i�Ⴆ��1�s�ځj���������āA�Ή������ԍ����擾
    Set headerRange = headerSourceSheet.Rows(1)

    ' ���x�����Ɋ�Â��ė�ԍ����������ACollection�Ɋi�[
    For Each Label In labelNames
        Set cell = headerRange.Find(Label, LookIn:=xlValues, LookAt:=xlWhole)
        If Not cell Is Nothing Then
            colNum = cell.Column
            columnsToKeep.Add colNum ' ����������ԍ���ǉ�
        Else
            MsgBox "���x�� " & Label & " ��������܂���B"
        End If
    Next Label

    ' �V�[�g���ύX�������Ăяo��
    Call RenameSheetsBasedOnA0No

    ' Sheet1�Ȃǂ̃V�[�g�ȊO�����[�v���ď���
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Sheet1" And ws.Name <> "�S�̃t���[" And ws.Name <> "�菇����" And ws.Name <> "�����" And ws.Name <> "�����" And ws.Name <> "Innovator" And ws.Name <> "���{" And ws.Name <> "Innovator (2)" Then
            ' �����t���[
            Call ProcessSheetColumns(ws, columnsToKeep)
            Call ProcessAdditionalColumns(ws)
            Call FormatAndStyleSheet(ws)
            Call InsertFormulasAndFormatting(ws, yaruyaraSheet)
            Call CopyDataToYaruyara(ws, yaruyaraSheet)
            Call RestrictInputBasedOnColumns(yaruyaraSheet)
        End If
    Next ws

    ' �����V�[�g�̎d�グ
    Call SetupYaruyaraSheet(yaruyaraSheet, headerSourceSheet)
End Sub

Sub RenameSheetsBasedOnA0No()
    Dim ws As Worksheet
    Dim cellValue As Variant
    Dim newName As String
    Dim colA0No As Variant
            
    ' �V�[�g�����[�v���ăV�[�g����ύX
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Sheet1" And ws.Name <> "�S�̃t���[" And ws.Name <> "�菇����" And ws.Name <> "�����" And ws.Name <> "�����" And ws.Name <> "Innovator" And ws.Name <> "���{" And ws.Name <> "Innovator (2)" Then

            ' A0 No.�������
            colA0No = Application.Match("A0 No.", ws.Rows(1), 0)
            
            ' A0 No.�����������ꍇ
            If Not IsError(colA0No) Then
                ' A0 No.���2�s�ڂ̒l���m�F
                cellValue = ws.Cells(2, colA0No).Value
            Else
                MsgBox """A0 No.""" & " ���x����������܂���ł����B"
            End If

            ' �V�[�g����ύX�i�G���[���������Ȃ��悤�ɏ����j
            On Error Resume Next
            ws.Name = newName
            On Error GoTo 0
        End If
    Next ws
End Sub

Sub ProcessSheetColumns(ws As Worksheet, columnsToKeep As Variant) '�s�v��폜�E�v���ԍ���}��
    Dim col As Long
    Dim colReqNo As Variant
    Dim colA0No As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim colItem As Variant
    Dim found As Boolean
    
    For col = ws.UsedRange.Columns.Count To 1 Step -1
        found = False
        ' columnsToKeep �̂��ׂĂ̒l�Ɣ�r
        For Each colItem In columnsToKeep
            If ws.Cells(1, col).Value = colItem Then
                found = True
                Exit For
            End If
        Next colItem
    
        ' ������Ȃ������ꍇ�A����폜
        If Not found Then ws.Columns(col).Delete
    Next col
    
    ' A���}��
    ws.Columns("A").Insert Shift:=xlToRight
    ws.Columns("A").NumberFormat = "@"

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' �V�[�g���𕶎���Ƃ��ē���
    For i = 2 To lastRow
        ws.Cells(i, 1).Value = "'" & ws.Name
    Next i

    ' �w�b�_�[��ݒ�
    ws.Cells(1, 1).Value = "�v���ԍ�"

    ' �G���[�`�F�b�N�𖳌����i�K�v�ɉ����āj
    Application.ErrorCheckingOptions.NumberAsText = False

    ' �v���ԍ��������
    colReqNo = Application.Match("�v���ԍ�", ws.Rows(1), 0)

    If Not IsError(colReqNo) Then
        ' A0 No.�������
        colA0No = Application.Match("A0 No.", ws.Rows(1), 0)
        
        If Not IsError(colA0No) Then
            ' A0 No.���v���ԍ���̉E�ׂɃR�s�[���đ}��
            ws.Columns(colA0No).Copy
            ws.Columns(colReqNo + 1).Insert Shift:=xlToRight
            
            ' �R�s�[��A����A0 No.����폜
            ws.Columns(colA0No + 1).Delete Shift:=xlToLeft
        Else
            MsgBox """A0 No.""" & " ���x����������܂���ł����B"
        End If
    Else
        MsgBox """�v���ԍ�""" & " ���x����������܂���ł����B"
    End If
End Sub

Sub ProcessAdditionalColumns(ws As Worksheet) '��}���E���ۗ�ړ�
    Dim colAheadInput As Variant
    Dim colRoom As Variant

    ' �uAHEAD���͉ہv��̔ԍ�������
    ws.Cells(1, 23).Value = "AHEAD���͉�"
    colAheadInput = Application.Match("AHEAD���͉�", ws.Rows(1), 0)

    If Not IsError(colAheadInput) Then
        ' �u���ہv�������
        colRoom = Application.Match("����", ws.Rows(1), 0)

        If Not IsError(colRoom) Then
            ' ���ۗ���E�ׂɈړ��i�}���j
            ws.Columns(colRoom).Cut
            ws.Columns(colAheadInput + 1).Insert Shift:=xlToRight
        Else
            MsgBox """����""" & " ���x����������܂���ł����B"
        End If
    Else
        MsgBox """AHEAD���͉�""" & " ���x����������܂���ł����B"
    End If
    
       ' �̔ۃ}�[�N��̍����Ɏ���5��}��
    ws.Columns("F:J").Insert Shift:=xlToRight
    ws.Range("F2:J" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).Value = ""

    ' �̔۔��藝�R��̍����Ɏ���5��}��
    ws.Columns("M:Q").Insert Shift:=xlToRight
    ws.Range("M2:Q" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row).Value = ""

    ' �}����������O���[�v��
    ws.Columns("F:J").Group
    ws.Columns("M:Q").Group
End Sub

Sub FormatAndStyleSheet(ws As Worksheet)
    Dim colorToApply As Long
    Dim whereESS As Long
    Dim lastRow As Long
    Dim cell As Range

    ' �V�[�g��1�s�ڂɃ��x�����i�[
    Dim headers As Variant
    headers = Array("BAT���\", "QJB MJB", "�\��", "ESS�M�}�l", "BTS�M�}�l", "", "BAT���\", "QJB MJB", "�\��", "ESS�M�}�l", "BTS�M�}�l", "�̔۔��藝�R", "�̔ۗ��R�`�F�b�N (�ϑ���)", "�̔ۗ��R�`�F�b�N (�ϑ���)", "���F�A�C�e��A�v���t�\", "���胉���N", "", "", "�@��S������", "����v��")
    
    Dim i As Integer
    For i = LBound(headers) To UBound(headers)
        If headers(i) <> "" Then
            ws.Cells(1, i + 6).Value = headers(i)
        End If
    Next i

    ' 1�s�ڂŐF�������Ă���Z�����������A���̐F�𑼂̃Z���ɓK�p
    For Each cell In ws.Rows(1).Cells
        If cell.Interior.ColorIndex <> -4142 Then
            colorToApply = cell.Interior.Color ' �F���擾
            Exit For
        End If
    Next cell

    ' �F�𑼂�1�s�ڂ̃Z���ɓK�p
    For Each cell In ws.Rows(1).Cells
        If cell.Interior.ColorIndex = -4142 Then
            cell.Interior.Color = colorToApply
        End If
    Next cell

    ' ���ۑ��œ��͂��������Ă��Ȃ��Z���ɉ��F��t����
    whereESS = 23
    lastRow = ws.Cells(ws.Rows.Count, whereESS).End(xlUp).Row
    For Each cell In ws.Range(ws.Cells(2, whereESS), ws.Cells(lastRow, whereESS))
        If Trim(cell.Value) <> "" And InStr(UCase(Trim(cell.Value)), "ESS") = 0 Then
            cell.EntireRow.Interior.Color = RGB(169, 169, 169)
        End If
    Next cell

    ' ������1�s�ڂ����F�œh��Ԃ�
    ws.Range("F1:J1").Interior.Color = RGB(255, 255, 0)
    ws.Range("L1:Q1").Interior.Color = RGB(255, 255, 0)
    ws.Range("T1:U1").Interior.Color = RGB(255, 255, 0)

    ' �t�B���^�[��ݒ�
    ws.Rows(1).AutoFilter

    ' �񕝂���������
    ws.Columns.AutoFit

    ' ���ׂẴV�[�g��A�񂩂�Y��܂ő����i�q������
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

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "�f�[�^������܂���B", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' �V�[�g�ی����
    ws.Unprotect

    ' �Čv�Z������
    Application.Calculate

    ' K��ɐ�����ݒ� (F�񂩂�J��ɕύX���������ꍇ�ł��Čv�Z)
    For Each cell In ws.Range("K2:K" & lastRow)
        ' ��ɐ������Đݒ�
        cell.FormulaR1C1 = _
           "=IF(COUNTA(RC[-5]:RC[-1])=0,"""",IF(COUNTIF(RC[-5]:RC[-1],""�Z"")>0,""�Z"",IF(COUNTIF(RC[-5]:RC[-1],""�~"")>0,""�~"",""-"")))"
    Next cell

    ' Z��ɐ�����ǉ�
    For Each cell In ws.Range("Z2:Z" & lastRow)
        If Len(Trim(cell.Value)) = 0 Or Not cell.HasFormula Then
            cell.Formula = "=IF(OR(K" & cell.Row & "=""-"", K" & cell.Row & "=""""), ""�S�ĊY������"", IF(K" & cell.Row & "=""�~"", ""�S�ăe�X�g�E�m�F����"", IF(K" & cell.Row & "=""�Z"", ""�e�X�g�v"", """")))"
        End If
    Next cell

    ' Y��ɐ�����ǉ�
    For Each cell In ws.Range("Y2:Y" & lastRow)
        If Len(Trim(cell.Value)) = 0 Or Not cell.HasFormula Then
            cell.Formula = "=IFERROR(IF(Z" & cell.Row & "=""�e�X�g�v"", ""�e�X�g�v"", IF(Z" & cell.Row & "=""�S�ăe�X�g�E�m�F����"", IF(IFERROR(COUNTIFS(W$1:W$10001, W" & cell.Row & ", Z$1:Z$10001, ""�e�X�g�v""), 0)>0, ""�e�X�g�v"", ""�S�ăe�X�g�E�m�F����""), IF(Z" & cell.Row & "=""�S�ĊY������"", IF(IFERROR(COUNTIFS(W$1:W$10001, W" & cell.Row & ", Z$1:Z$10001, ""�e�X�g�v""), 0)>0, ""�e�X�g�v"", IF(IFERROR(COUNTIFS(W$1:W$10001, W" & cell.Row & ", Z$1:Z$10001, ""�S�ăe�X�g�E�m�F����""), 0)>0, ""�S�ăe�X�g�E�m�F����"", ""�S�ĊY������"")), Z" & cell.Row & "))), """")"
        End If
    Next cell
    
    ' V��ɏ����t��������ǉ�
    Set targetRange = ws.Range("A2:Z" & lastRow)
    With targetRange
        .FormatConditions.Delete ' �����̏����t���������폜
        ' �����𗘗p���������t��������ǉ�
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$V2=""�Z""")
            .Interior.Color = RGB(255, 255, 255) ' ���F�ɐݒ�
        End With
    End With
        
    ' Z����\��
    ws.Columns("Z").Hidden = True
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical
    On Error GoTo 0
End Sub

Sub CopyDataToYaruyara(ws As Worksheet, yaruyaraSheet As Worksheet) '�����V�[�g�ւ̃R�s�y����
    Dim yaruyaraLastRow As Long
    Dim lastRow As Long
    Dim dataRange As Range
    Dim destRange As Range
    
    ' �f�[�^�̍ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 2�s�ڈȍ~�̃f�[�^�͈͂�ݒ�
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.Columns.Count))

    ' �����V�[�g�̍ŏI�s���擾
    yaruyaraLastRow = yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row

    ' �\��t����͈̔͂�ݒ�
    Set destRange = yaruyaraSheet.Cells(yaruyaraLastRow + 1, 1).Resize(dataRange.Rows.Count, dataRange.Columns.Count)

    ' �f�[�^�ƃt�H�[�}�b�g���R�s�[
    dataRange.Copy
    destRange.PasteSpecial Paste:=xlPasteAll

    ' �R�s�[��ɃN���b�v�{�[�h���N���A
    Application.CutCopyMode = False
    
    ' �����V�[�g�ł�����������K�p
    destRange.HorizontalAlignment = xlCenter
End Sub

Sub SetupYaruyaraSheet(yaruyaraSheet As Worksheet, headerSourceSheet As Worksheet)
    Dim lastRow As Long
    Dim targetRange As Range
    
    ' �ی������
    If yaruyaraSheet.ProtectContents Then
        yaruyaraSheet.Unprotect Password:="password" ' �p�X���[�h�͓K�X�ݒ�
    End If

    ' �w�b�_�[1�s�ڂ������V�[�g�ɃR�s�[
    headerSourceSheet.Rows(1).Copy Destination:=yaruyaraSheet.Rows(1)

    ' �����V�[�g��1�s�ڂɃt�B���^�[��ݒ�
    yaruyaraSheet.Rows(1).AutoFilter

    ' A�񂩂�Y��܂ő����i�q������
    With yaruyaraSheet.Range("A1:Y" & yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = 0
    End With

    ' �}����������O���[�v��
    yaruyaraSheet.Columns("F:J").Group
    yaruyaraSheet.Columns("L:P").Group

    ' �ŏI�s���擾
    lastRow = yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row

    ' V��̏����t��������ݒ�
    Set targetRange = yaruyaraSheet.Range("A2:Z" & lastRow)
    With targetRange
        .FormatConditions.Delete ' �����̏����t���������폜
        ' �����𗘗p���������t��������ǉ�
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$V2=""�Z""")
            .Interior.Color = RGB(255, 255, 255) ' ���F�ɐݒ�
        End With
    End With

    ' �Z���̓��e�𐅕������ɒ�������
    yaruyaraSheet.Range("A1:Y" & yaruyaraSheet.Cells(yaruyaraSheet.Rows.Count, "A").End(xlUp).Row).HorizontalAlignment = xlCenter

    ' �񕝂���������
    yaruyaraSheet.Columns.AutoFit

    ' Z����\��
    yaruyaraSheet.Columns("Z").Hidden = True

    ' A��������Ƀ\�[�g
    With yaruyaraSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=yaruyaraSheet.Range("A2:A" & lastRow), Order:=xlAscending
        .SetRange yaruyaraSheet.Range("A1:Y" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' �ی���Đݒ�
    yaruyaraSheet.Protect Password:="password", UserInterfaceOnly:=True

    MsgBox "�������������܂����I"
End Sub

Sub RestrictInputBasedOnColumns(ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim qFilled As Boolean
    Dim kFilled As Boolean

    ' �ی����
    If ws.ProtectContents Then
        ws.Unprotect Password:="password"
    End If

    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' ���͐�����ݒ�
    For Each cell In ws.Range("A2:Z" & lastRow).Cells
        If Not ((cell.Column >= 6 And cell.Column <= 10) Or _
                (cell.Column >= 12 And cell.Column <= 17) Or _
                (cell.Column >= 20 And cell.Column <= 21)) Then
            cell.Locked = True
        Else
            cell.Locked = False
        End If
    Next cell

    ' Q���K�񂪖��܂����ꍇ�̃`�F�b�N
    qFilled = WorksheetFunction.CountBlank(ws.Range("Q2:Q" & lastRow)) = 0
    kFilled = WorksheetFunction.CountBlank(ws.Range("K2:K" & lastRow)) = 0

    If qFilled And kFilled Then
        ' Q���K�񂪖��܂�����V���X������
        ws.Range("V2:V" & lastRow).Locked = False
        ws.Range("X2:X" & lastRow).Locked = False
    Else
        ' �ă��b�N
        ws.Range("V2:V" & lastRow).Locked = True
        ws.Range("X2:X" & lastRow).Locked = True
    End If

    ' �V�[�g�ی���Đݒ�
    ws.Protect Password:="password", UserInterfaceOnly:=True, AllowFiltering:=True
End Sub


