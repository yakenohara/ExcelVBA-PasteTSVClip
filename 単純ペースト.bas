Attribute VB_Name = "�P���y�[�X�g"

Sub �P���y�[�X�g()

    Dim CB As New DataObject
    
    Dim delimiter As String: delimiter = vbTab '�Z���ԋ�؂蕶��
    Dim NLCharacter As String: NLCharacter = vbCrLf '���s����
    
    Dim maxNumOfLines As Long
    Dim maxNumOfColumns As Long
    
    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��Ēl�̏������݂��s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
    
    '���s�m�F
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    '�V�[�g�I����ԃ`�F�b�N
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    CB.GetFromClipboard '�N���b�v�{�[�h�̓��e���擾
    
    If Not (CB.GetFormat(1)) Then '�e�L�X�g�`���łȂ��ꍇ
        retVal = MsgBox("�N���b�v�{�[�h�Ƀe�L�X�g������܂���", vbExclamation)
        Exit Sub '�I��
        
    End If
    
    linesOfToPasteText = Split(CB.GetText, NLCharacter) '�s��؂�̕�����z����擾
    
    maxNumOfLines = UBound(linesOfToPasteText)
    maxNumOfColumns = 0
    
    '�㏑���m�F�p�ő�񐔂̌���
    For lineCounter = 0 To maxNumOfLines '�s���[�v for �ő�column������
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        numOfColumns = UBound(toPasteStrings)
        
        If (maxNumOfColumns < numOfColumns) Then
        
            maxNumOfColumns = numOfColumns '�ő�񐔂̕ۑ�
            
        End If
    
    Next lineCounter
    
    '�����ݐ�Z���͈͂̑I��
    Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row + maxNumOfLines, Selection.Column + maxNumOfColumns)).Select
    
    '�㏑���m�F
    If WorksheetFunction.CountA(Selection) > 0 Then
        yn = MsgBox("�y�[�X�g��̃Z���ɒl�������Ă��܂�" & vbLf & vbLf & _
                    "�㏑�����܂����H", _
                    vbOKCancel)
        
        If yn = vbCancel Then
            Exit Sub
        End If
    
    End If
    
    '�y�[�X�g���[�v
    For lineCounter = 0 To maxNumOfLines '�s���[�v
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        
        numOfColumns = UBound(toPasteStrings)
        
        For columnCounter = 0 To numOfColumns '�񃋁[�v
            
            '�����Z���̏ꍇ�́A�����Z���̍���ɂ̂݃y�[�X�g����
            If (Selection(1).Offset(lineCounter, columnCounter).Address = Selection(1).Offset(lineCounter, columnCounter).MergeArea.Cells(1, 1).Address) Then
                
                Selection(1).Offset(lineCounter, columnCounter).Value = toPasteStrings(columnCounter)
                
            End If
            
        Next columnCounter
    
    Next lineCounter
    
End Sub

