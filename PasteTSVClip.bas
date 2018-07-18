Attribute VB_Name = "PasteTSVClip"
Sub PasteTSVClip()

    Dim CB As New DataObject
    
    Dim delimiter As String: delimiter = vbTab '�Z���ԋ�؂蕶��
    Dim NLCharacter As String: NLCharacter = vbCrLf '���s����
    
    Dim maxNumOfLines As Long
    Dim maxNumOfColumns As Long
    
    Dim ignoreVacant As Boolean '�󕶎��𖳎����邩�ǂ���
    
    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��Ēl�̏������݂��s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
                                                   
    Dim noClipMessage As String: noClipMessage = "�N���b�v�{�[�h�Ƀe�L�X�g������܂���"
    
    Dim conflictMessage As String: conflictMessage = "�y�[�X�g��̃Z���ɒl�������Ă��܂�" & vbLf & vbLf & _
                                                     "�u�Ď��s�v�ŏ㏑��" & vbLf & _
                                                     "�u�����v�Ńy�[�X�g���̋󔒂̂ݖ������ăy�[�X�g���܂�"
    
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
        retVal = MsgBox(noClipMessage, vbExclamation)
        Exit Sub '�I��
        
    End If
    
    linesOfToPasteText = Split(CB.GetText, NLCharacter) '�s��؂�̕�����z����擾
    
    maxNumOfLines = UBoundSafe(linesOfToPasteText)
    
    '�N���b�v�{�[�h���󂾂����ꍇ
    If maxNumOfLines < 0 Then
        retVal = MsgBox(noClipMessage, vbExclamation)
        Exit Sub '�I��
    
    End If
    
    '�ŏI�s���󕶎��������ꍇ�́A���̍s�͏������Ȃ�
    If linesOfToPasteText(maxNumOfLines) = "" Then
        maxNumOfLines = maxNumOfLines - 1
    End If

    '�㏑���m�F�p�ő�񐔂̌���
    maxNumOfColumns = 0
    For lineCounter = 0 To maxNumOfLines '�s���[�v for �ő�column������
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        numOfColumns = UBoundSafe(toPasteStrings)
        
        If (maxNumOfColumns < numOfColumns) Then
        
            maxNumOfColumns = numOfColumns '�ő�񐔂̕ۑ�
            
        End If
    
    Next lineCounter
    
    '�����ݐ�Z���͈͂̑I��
    Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row + maxNumOfLines, Selection.Column + maxNumOfColumns)).Select
    
    ignoreVacant = True '�f�t�H���g�́A�󕶎��𖳎�����
    
    '�㏑���m�F
    If WorksheetFunction.CountA(Selection) > 0 Then
        ri = MsgBox(conflictMessage, vbAbortRetryIgnore)
        
        If ri = vbAbort Then '[���~]�{�^���������ꂽ
            Exit Sub
        
        ElseIf ri = vbRetry Then '[�Ď��s]�{�^���������ꂽ
            ignoreVacant = False '�󕶎����㏑������
            
        Else '[����]�{�^���������ꂽ
            ignoreVacant = True '�󕶎��͖�������
            
        End If
    
    End If
    
    '�y�[�X�g���[�v
    For lineCounter = 0 To maxNumOfLines '�s���[�v
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        numOfColumns = UBoundSafe(toPasteStrings)
        
        If numOfColumns < 0 Then '��z�񂾂����ꍇ
            toPasteStrings = Array("") '�󕶎�1�̔z����`����
            numOfColumns = 0

        End If
        
        For columnCounter = 0 To numOfColumns '�񃋁[�v
            
            '�����Z���̏ꍇ�́A�����Z���̍���ɂ̂݃y�[�X�g����
            If (Selection(1).Offset(lineCounter, columnCounter).Address = Selection(1).Offset(lineCounter, columnCounter).MergeArea.Cells(1, 1).Address) Then
                
                '�󕶎��̏ꍇ�́A�����ݒ肳��Ă��Ȃ���Γ\��t����
                If (toPasteStrings(columnCounter) <> "") Or _
                   ((toPasteStrings(columnCounter) = "") And Not (ignoreVacant)) Then
                   
                    Selection(1).Offset(lineCounter, columnCounter).Value = toPasteStrings(columnCounter)
                    
                End If
                
            End If
            
        Next columnCounter
    
    Next lineCounter
    
End Sub

Private Function UBoundSafe(ar As Variant) As Long
    Dim tmp As Long
    
On Error GoTo ERROR_

    tmp = UBound(ar)
    
    UBoundSafe = tmp
        
    Exit Function

ERROR_:
    tmp = -1
    Resume Next
    
End Function


