Attribute VB_Name = "�C"
Sub finalCheckMacro()
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
    Dim A0NoCol As Long
    Dim A0NoValue As String

    ' �u�����v�V�[�g���m�F���Đݒ�
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("�����")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "�u�����v�V�[�g��������܂���B�������I�����܂��B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g�����X�g�̍쐬
    sheetList = "�V�[�g�����X�g:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    ' ���[�U�[�ɃV�[�g�ԍ�����͂�����
    sheetIndex = Application.InputBox("��r����V�[�g��I�����Ă��������i�ԍ�����́j:" & vbCrLf & sheetList, "�V�[�g�I��", Type:=1)

    ' ���̓`�F�b�N
    If sheetIndex < 1 Or sheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "�������V�[�g�ԍ�����͂��Ă��������B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g2��ݒ�
    Set Sheet2 = ThisWorkbook.Sheets(sheetIndex)
    
    ' �uA0 No.�v���x���̗�ԍ����擾
    A0NoCol = Application.Match("A0 No.", Sheet2.Rows(1), 0)
    
    ' �uA0 No.�v�����������ꍇ�A���̗��2�s�ڂ̒l�̍�����4�������擾
    If Not IsError(A0NoCol) Then
        A0NoValue = Left(Sheet2.Cells(2, A0NoCol).Value, 4)
    Else
        MsgBox """A0 No.""" & " ���x����������܂���ł����B"
        Exit Sub
    End If
    
    ' keyValue�Ɋi�[
    keyValue = Trim(A0NoValue)

    ' �V�[�g1��A��ŃL�[�l���n�܂�s������
    startRow1 = Sheet1.Columns(1).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole).Row

    ' �V�[�g1��A��ŃL�[�l���I���s������
    endRow1 = Sheet1.Columns(1).Find(What:=keyValue & "*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row

    ' �G���[����
    If startRow1 = 0 Or endRow1 = 0 Then
        MsgBox "�L�[�l�i" & keyValue & "�j��������܂���ł����B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g1�i�����j�Ŕ�r������I��
    Set range1 = Application.InputBox("��r�����������V�[�g�őI�����Ă��������i��: =�����!$B:$B�j�B", Type:=8)
    If range1 Is Nothing Or range1.Worksheet.Name <> "�����" Then
        MsgBox "�����V�[�g�̗񂪐������I������Ă��܂���B�������I�����܂��B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g2�Ŕ�r������I��
    Set range2 = Application.InputBox("��r�����𑼂̃V�[�g�őI�����Ă��������i��: =" & Sheet2.Name & "!$P:$P�j�B", Type:=8)
    If range2 Is Nothing Then
        MsgBox "�V�[�g2�̗񂪑I������܂���ł����B�������I�����܂��B", vbExclamation
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

        ' �l���擾���Ĕ�r
        val1 = Trim(CStr(cell1.Value))
        val2 = Trim(CStr(cell2.Value))

        ' ��r����
        If val1 <> val2 Then
            cell1.Interior.Color = RGB(255, 0, 0)
            cell2.Interior.Color = RGB(255, 0, 0)
            mismatchDetails = mismatchDetails & "�V�[�g1�s " & row1 & " / �V�[�g2�s " & row2 & ": �l����v���܂��� (Cell1: [" & val1 & "], Cell2: [" & val2 & "])" & vbCrLf
        End If

        ' �V�[�g2�̎��̍s��
        row2 = row2 + 1
    Next row1

    ' ���ʂ�\��
    If mismatchDetails = "" Then
        MsgBox "���ׂĈ�v���܂����I", vbInformation
    Else
        MsgBox "�ȉ��̕s��v��������܂���:" & vbCrLf & mismatchDetails, vbExclamation
        ' �s��v�s��V�����V�[�g�ɏ����o��
        WriteMismatchToNewSheet mismatchDetails
    End If
End Sub

Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' �V�����V�[�g��ǉ�
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "�s��v�s(�ŏI�`�F�b�N�}�N��)"

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





