Attribute VB_Name = "�B"
Sub compareyaruyara_hanteiyouhi() '����,����v�ۗ�̊m�F
    Dim requirementNumber As String ' �v���ԍ����i�[���邽�߂̕�����ϐ�
    Dim requirementRow As Long ' �v���ԍ������������s�ԍ����i�[���邽�߂̕ϐ�
    Dim roomKaa As String ' ���ہi�����炭�����S���j�̏����i�[���邽�߂̕�����ϐ�
    Dim yColumnValue As String ' Y��̒l���i�[���邽�߂̕�����ϐ�
    Dim roomColumn As Long ' ���ۗ�̗�ԍ����i�[���邽�߂̕ϐ�
    Dim judgementColumn As Long ' �����̗�ԍ����i�[���邽�߂̕ϐ�
    Dim lastRow1 As Long ' �ŏ��̃V�[�g�iSheet1�j�ɂ�����f�[�^�̍ŏI�s���i�[����ϐ�
    Dim lastRow2 As Long ' 2�Ԗڂ̃V�[�g�iSheet2�j�ɂ�����f�[�^�̍ŏI�s���i�[����ϐ�
    Dim matchResult As Variant ' Match�֐��̌��ʁi�������ꂽ�ʒu�j���i�[���邽�߂̕ϐ�
    Dim rowNumber As Long ' �s�ԍ����i�[���邽�߂̕ϐ�
    Dim Sheet1 As Worksheet ' �ŏ��̃V�[�g�iSheet1�j���i�[���邽�߂�Worksheet�I�u�W�F�N�g
    Dim Sheet2 As Worksheet ' 2�Ԗڂ̃V�[�g�iSheet2�j���i�[���邽�߂�Worksheet�I�u�W�F�N�g
    Dim MismatchRows As String ' �s��v�s���i�[���邽�߂̕�����ϐ�
    Dim selectedSheetIndex As Long

    MismatchRows = "" ' ������

    ' �V�[�g1���u�����v�Ƃ������O�ŌŒ�
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("�����")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "�����V�[�g������܂���B", vbExclamation
        Exit Sub
    End If

    ' ���[�U�[�ɗv���ԍ���I��������
    lastRow1 = Sheet1.Cells(Sheet1.Rows.Count, "A").End(xlUp).Row
    Dim requirementList As String
    Dim i As Long
    For i = 2 To lastRow1
        requirementList = requirementList & Sheet1.Cells(i, "A").Value & vbCrLf
    Next i

    requirementNumber = InputBox("�ȉ��̗v���ԍ�����I�����Ă�������:" & vbCrLf & requirementList, "�v���ԍ��I��")

    ' �v���ԍ������͂���Ă��Ȃ��ꍇ�͏I��
    If Len(Trim(requirementNumber)) = 0 Then
        MsgBox "�v���ԍ������͂���Ă��܂���B", vbExclamation
        Exit Sub
    End If

    ' �V�[�g1��A����������Ĉ�v����v���ԍ���T��
    matchResult = Application.Match(requirementNumber, Sheet1.Range("A2:A" & lastRow1), 0)

    ' ��v���Ȃ��ꍇ�A�G���[���b�Z�[�W��\�����ďI��
    If IsError(matchResult) Then
        MsgBox "�w�肳�ꂽ�v���ԍ��̓V�[�g1��A��ɑ��݂��܂���B", vbExclamation
        Exit Sub
    End If

    ' ��v�����ꍇ�A���̍s�̔ԍ����擾
    requirementRow = matchResult + 1 ' matchResult��1����n�܂邽�ߒ���

    ' �V�[�g�����X�g�̍쐬
    Dim sheetList As String
    sheetList = "�V�[�g�����X�g:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    ' �V�[�g2��I���i�ԍ��őI�ԁj
    selectedSheetIndex = CInt(InputBox("��r�������V�[�g��I�����Ă�������:" & vbCrLf & sheetList))
    If selectedSheetIndex < 1 Or selectedSheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "�����Ȕԍ��ł��B"
        Exit Sub
    End If
    Set Sheet2 = ThisWorkbook.Sheets(selectedSheetIndex)

    ' �V�[�g1��1�s�ڂ���u���ہv�Ɓu����v�ہv�̗������
    roomColumn = Application.Match("����", Sheet1.Rows(1), 0)
    judgementColumn = Application.Match("����v��", Sheet1.Rows(1), 0)

    ' �u���ہv��Ɓu����v�ہv�񂪌�����Ȃ��ꍇ�A�G���[���o�͂��ďI��
    If IsError(roomColumn) Then
        MsgBox """����""�񂪌�����܂���ł����B"
        Exit Sub
    End If
    If IsError(judgementColumn) Then
        MsgBox """����v��""�񂪌�����܂���ł����B"
        Exit Sub
    End If

    ' �V�[�g2�̍ŏI�s���擾
    lastRow2 = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row

    ' �V�[�g2��2�s�ڂ���ŏI�s�܂Ń��[�v
    For rowNumber = 2 To lastRow2
        ' �V�[�g2��A��i���ہj���擾
        roomKaa = Sheet2.Cells(rowNumber, 1).Value

        ' �V�[�g1�̈�v����v���ԍ��̍s�Łu���ہv�̒l���擾
        If Sheet1.Cells(requirementRow, roomColumn).Value = roomKaa Then
            ' ���ۂ���v�����ꍇ�A����v�ۗ�̒l���擾
            yColumnValue = Sheet1.Cells(requirementRow, judgementColumn).Value

            ' �V�[�g2��C��ƈ�v���Ă��邩�m�F
            If Trim(yColumnValue) = "" And Trim(Sheet2.Cells(rowNumber, 3).Value) = "" Then
                ' �������󔒂̏ꍇ�͉������Ȃ�
            ElseIf yColumnValue = Sheet2.Cells(rowNumber, 3).Value Then
                ' ��v����ꍇ�͉������Ȃ�
            Else
                ' �s��v�̏ꍇ�A�V�[�g1�̔���v�ۗ�ƃV�[�g2��C���ԐF�œh��Ԃ�
                Sheet1.Cells(requirementRow, judgementColumn).Interior.Color = RGB(255, 0, 0) ' �V�[�g1�̔���v�ۗ��ԐF
                Sheet2.Cells(rowNumber, 3).Interior.Color = RGB(255, 0, 0) ' �V�[�g2��C���ԐF
                ' �s��v�s�̏������W
                MismatchRows = MismatchRows & "�V�[�g1�s" & requirementRow & "�ƃV�[�g2�s" & rowNumber & vbCrLf
            End If
        End If
    Next rowNumber

    ' �s��v�s�����݂���ꍇ�A�V�����V�[�g�Ƀ������쐬
    If MismatchRows <> "" Then
        Call WriteMismatchToNewSheet(MismatchRows)
    Else
        MsgBox "�s��v�͌�����܂���ł����B"
    End If
End Sub

Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' �V�����V�[�g��ǉ�
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "�s��v�s�i����v�ہj"

    ' �w�b�_�[����������
    NewSheet.Cells(1, 1).Value = "�s��v�s�̏ڍ�"

    ' mismatchRows �����s�ŕ������Ĕz��Ɋi�[
    Lines = Split(MismatchRows, vbCrLf)

    ' �s��v������������
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex

    ' �����ʒm
    MsgBox "�s��v�s�̏ڍׂ��V�����V�[�g�ɕۑ�����܂����B"
End Sub


   
