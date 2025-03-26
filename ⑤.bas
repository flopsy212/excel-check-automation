Attribute VB_Name = "�D"
Sub finalLastcheckfiles_macrosheet() ' ����҃V�[�g�̖�����Email�񂾂����͂���AHEAD�ƍ��v���邩�m�F
    Dim Sheet1 As Worksheet
    Dim Sheet2 As Worksheet
    Dim lastRow1 As Long
    Dim lastRow2 As Long
    Dim rowIdx As Long
    Dim RoleColumn2 As Long
    Dim EmailColumn2 As Long
    Dim mismatchDetails As String

    ' �V�[�g1��ݒ�
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("�����")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "�����V�[�g������܂���B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g2��I��
    Dim selectedSheetIndex As Integer
    Dim SheetNameList As String
    SheetNameList = "�V�[�g�����X�g:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        SheetNameList = SheetNameList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    selectedSheetIndex = CInt(InputBox("��r�������V�[�g��I�����Ă��������B" & vbCrLf & SheetNameList))
    If selectedSheetIndex < 1 Or selectedSheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "�����Ȕԍ��ł��B"
        Exit Sub
    End If

    Set Sheet2 = ThisWorkbook.Sheets(selectedSheetIndex)

    ' �V�[�g1�ƃV�[�g2�̍ŏI�s���擾
    lastRow1 = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row

    ' �V�[�g2�́u�����v�ƁuEmail�v������
    RoleColumn2 = 0
    EmailColumn2 = 0
    For ColIdx = 1 To Sheet2.Cells(1, Columns.Count).End(xlToLeft).Column
        If Trim(Sheet2.Cells(1, ColIdx).Value) = "����" Then
            RoleColumn2 = ColIdx
        ElseIf Trim(Sheet2.Cells(1, ColIdx).Value) = "Email" Then
            EmailColumn2 = ColIdx
        End If
    Next ColIdx

    ' �G���[����
    If RoleColumn2 = 0 Or EmailColumn2 = 0 Then
        MsgBox "�V�[�g2�ɕK�v�ȗ�i�����܂���Email�j��������܂���B", vbExclamation
        Exit Sub
    End If

    ' ��r�J�n
    mismatchDetails = ""
    For rowIdx = 2 To Application.WorksheetFunction.Min(lastRow1, lastRow2)
        Dim Role1 As String, Email1 As String
        Dim Role2 As String, Email2 As String

        ' �V�[�g1�ƃV�[�g2�̒l���擾
        Role1 = Trim(Sheet1.Cells(rowIdx, 1).Value) ' �V�[�g1�̖��� (A��)
        Email1 = Trim(Sheet1.Cells(rowIdx, 2).Value) ' �V�[�g1��Email (B��)
        Role2 = Trim(Sheet2.Cells(rowIdx, RoleColumn2).Value) ' �V�[�g2�̖���
        Email2 = Trim(Sheet2.Cells(rowIdx, EmailColumn2).Value) ' �V�[�g2��Email

        ' ��r���ĕs��v���L�^
        If Role1 <> Role2 Or Email1 <> Email2 Then
            mismatchDetails = mismatchDetails & "�s " & rowIdx & " �ɕs��v������܂�:" & vbCrLf & _
                             "  �V�[�g1 - ����: " & Role1 & ", Email: " & Email1 & vbCrLf & _
                             "  �V�[�g2 - ����: " & Role2 & ", Email: " & Email2 & vbCrLf & vbCrLf
        End If
    Next rowIdx

    ' ���ʂ�V�����V�[�g�ɏo��
    If mismatchDetails <> "" Then
        Call WriteMismatchToNewSheet(mismatchDetails)
        MsgBox "�s��v��������܂����B�ڍׂ͐V�����V�[�g���m�F���Ă��������B", vbInformation
    Else
        MsgBox "�S�Ĉ�v���Ă��܂��B", vbInformation
    End If
End Sub

Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' �V�����V�[�g��ǉ�
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "�s��v�s�i�t�@�C�i���ŏI�j"

    ' �w�b�_�[����������
    NewSheet.Cells(1, 1).Value = "�s��v�s�̏ڍ�"

    ' MismatchRows �����s�ŕ������Ĕz��Ɋi�[
    Lines = Split(MismatchRows, vbCrLf)

    ' �s��v������������
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            ' ���̕���������̂܂܏�������
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub


