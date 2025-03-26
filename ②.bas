Attribute VB_Name = "�A"
Sub checkMacro()
    Dim range1 As Range, range2 As Range
    Dim cell1 As Range, cell2 As Range
    Dim mismatchDetails As String
    Dim val1 As Variant, val2 As Variant
    Dim row1 As Long, row2 As Long
    Dim startRow1 As Long, endRow1 As Long
    Dim keyValue As String
    Dim Sheet1 As Worksheet, Sheet2 As Worksheet
    Dim sheetList As String, sheetIndex1 As Integer, sheetIndex2 As Integer
    Dim A0NoCol As Long
    Dim A0NoValue As String

    ' �V�[�g�����X�g�̍쐬
    sheetList = "�V�[�g�����X�g:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    ' ���[�U�[�ɔ�r����V�[�g1��I��������
    sheetIndex1 = Application.InputBox("��r���̃V�[�g��I�����Ă��������i�ԍ�����́j:" & vbCrLf & sheetList, "�V�[�g�I��", Type:=1)
    If sheetIndex1 < 1 Or sheetIndex1 > ThisWorkbook.Sheets.Count Then
        MsgBox "�������V�[�g�ԍ�����͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' ���[�U�[�ɔ�r����V�[�g2��I��������
    sheetIndex2 = Application.InputBox("��r�Ώۂ̃V�[�g��I�����Ă��������i�ԍ�����́j:" & vbCrLf & sheetList, "�V�[�g�I��", Type:=1)
    If sheetIndex2 < 1 Or sheetIndex2 > ThisWorkbook.Sheets.Count Then
        MsgBox "�������V�[�g�ԍ�����͂��Ă��������B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g��ݒ�
    Set Sheet1 = ThisWorkbook.Sheets(sheetIndex1)
    Set Sheet2 = ThisWorkbook.Sheets(sheetIndex2)

    ' �uA0 No.�v���x���̗�ԍ����擾
    On Error Resume Next
    A0NoCol = Application.Match("A0 No.", Sheet2.Rows(1), 0)
    On Error GoTo 0
    
    ' �G���[����
    If A0NoCol = 0 Then
        MsgBox """A0 No.""" & " ���x����������܂���ł����B", vbExclamation
        Exit Sub
    End If

    ' �uA0 No.�v���2�s�ڂ̒l�̍�����4�������擾
    A0NoValue = Left(Sheet2.Cells(2, A0NoCol).Value, 4)
    keyValue = Trim(A0NoValue)

    ' �V�[�g1��A��ŃL�[�l���n�܂�s������
    Dim foundCell As Range
    Set foundCell = Sheet1.Columns(1).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        MsgBox "�L�[�l�i" & keyValue & "�j��������܂���ł����B", vbExclamation
        Exit Sub
    Else
        startRow1 = foundCell.Row
    End If

    ' �V�[�g1��A��ŃL�[�l���I���s������
    Set foundCell = Sheet1.Columns(1).Find(What:=keyValue & "*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        MsgBox "�L�[�l�i" & keyValue & "�j�͈̔͂�������܂���ł����B", vbExclamation
        Exit Sub
    Else
        endRow1 = foundCell.Row
    End If

    ' �V�[�g1�Ŕ�r������I��
    Set range1 = Application.InputBox("��r������I�����Ă��������i��: =" & Sheet1.Name & "!$K:$K�j�B", Type:=8)
    If range1 Is Nothing Or range1.Worksheet.Name <> Sheet1.Name Then
        MsgBox "��r���V�[�g�̗񂪐������I������Ă��܂���B�������I�����܂��B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g2�Ŕ�r������I��
    Set range2 = Application.InputBox("��r������I�����Ă��������i��: =" & Sheet2.Name & "!$AN:$AN�j�B", Type:=8)
    If range2 Is Nothing Then
        MsgBox "��r�ΏۃV�[�g�̗񂪑I������܂���ł����B�������I�����܂��B", vbExclamation
        Exit Sub
    End If

    ' ������
    mismatchDetails = ""

    ' ��r�����i1��1�̍s��r�j
    row2 = 2 ' �V�[�g2�̊J�n�s
    For row1 = startRow1 To endRow1
        If row2 > Sheet2.Cells(Sheet2.Rows.Count, range2.Column).End(xlUp).Row Then Exit For

        ' �Z�����擾
        Set cell1 = Sheet1.Cells(row1, range1.Column)
        Set cell2 = Sheet2.Cells(row2, range2.Column)

        ' �l���擾
        val1 = Trim(CStr(cell1.Value))
        val2 = Trim(CStr(cell2.Value))

        ' �w�i�F�ƒl�̔�r
    If cell1.Interior.Color = RGB(169, 169, 169) Or cell2.Interior.Color = RGB(169, 169, 169) Or _
   cell1.Interior.Color = RGB(166, 166, 166) Or cell2.Interior.Color = RGB(166, 166, 166) Then

        GoTo ContinueLoop
    Else
        ' ���h��Z���F���S��v���m�F
        If val1 <> val2 Then
            cell1.Interior.Color = RGB(255, 0, 0)
            cell2.Interior.Color = RGB(255, 0, 0)
            mismatchDetails = mismatchDetails & "�V�[�g1�s " & row1 & " / �V�[�g2�s " & row2 & ": ���h��Z���ŕs��v (Cell1: [" & val1 & "], Cell2: [" & val2 & "])" & vbCrLf
        End If
    End If
    
ContinueLoop:
    
        ' �V�[�g2�̎��̍s��
        row2 = row2 + 1
    Next row1

    ' ���ʂ�\��
    If mismatchDetails = "" Then
        MsgBox "���ׂĈ�v���܂����I", vbInformation
    Else
        MsgBox "�ȉ��̕s��v��������܂���:" & vbCrLf & mismatchDetails, vbExclamation
        ' �s��v�s��V�����V�[�g�ɏ����o��
        WriteMismatchResults mismatchDetails
    End If
End Sub

Sub WriteMismatchResults(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' �V�����V�[�g��ǉ�
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "�s��v�s(�`�F�b�N�}�N��)"

    ' �w�b�_�[����������
    NewSheet.Cells(1, 1).Value = "�s��v�s�̏ڍ�"

    ' MismatchRows �����s�ŕ������Ĕz��Ɋi�[
    Lines = Split(MismatchRows, vbCrLf)

    ' �s��v������������
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub





