Attribute VB_Name = "PasteTSVClip"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

Sub PasteTSVClip()

    Dim CBstr As String
    
    Dim delimiter As String: delimiter = vbTab 'セル間区切り文字
    Dim NLCharacter As String: NLCharacter = vbCrLf '改行文字
    
    Dim maxNumOfLines As Long
    Dim maxNumOfColumns As Long
    
    Dim ignoreVacant As Boolean '空文字を無視するかどうか
    
    Dim cautionMessage As String: cautionMessage = "このSubプロシージャは、" & vbLf & _
                                                   "現在の選択範囲に対して値の書き込みを行います。" & vbLf & vbLf & _
                                                   "実行しますか?"
                                                   
    Dim noClipMessage As String: noClipMessage = "クリップボードにテキストがありません"
    
    Dim conflictMessage As String: conflictMessage = "ペースト先のセルに値が入っています" & vbLf & vbLf & _
                                                     "「再試行」で上書き" & vbLf & _
                                                     "「無視」でペースト元の空白のみ無視してペーストします"
    
    '実行確認
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    'シート選択状態チェック
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "複数シートが選択されています" & vbLf & _
               "不要なシート選択を解除してください"
        Exit Sub
    End If
    
    GetCB CBstr 'クリップボードの内容を取得
    
    If (CBstr = "") Then 'テキスト形式でない場合
        retVal = MsgBox(noClipMessage, vbExclamation)
        Exit Sub '終了
        
    End If
    
    linesOfToPasteText = Split(CBstr, NLCharacter) '行区切りの文字列配列を取得
    
    maxNumOfLines = UBoundSafe(linesOfToPasteText)
    
    'クリップボードが空だった場合
    If maxNumOfLines < 0 Then
        retVal = MsgBox(noClipMessage, vbExclamation)
        Exit Sub '終了
    
    End If
    
    '最終行が空文字だった場合は、その行は処理しない
    If linesOfToPasteText(maxNumOfLines) = "" Then
        maxNumOfLines = maxNumOfLines - 1
    End If

    '上書き確認用最大列数の検査
    maxNumOfColumns = 0
    For lineCounter = 0 To maxNumOfLines '行ループ for 最大column数検査
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        numOfColumns = UBoundSafe(toPasteStrings)
        
        If (maxNumOfColumns < numOfColumns) Then
        
            maxNumOfColumns = numOfColumns '最大列数の保存
            
        End If
    
    Next lineCounter
    
    '書込み先セル範囲の選択
    Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row + maxNumOfLines, Selection.Column + maxNumOfColumns)).Select
    
    ignoreVacant = True 'デフォルトは、空文字を無視する
    
    '上書き確認
    If WorksheetFunction.CountA(Selection) > 0 Then
        ri = MsgBox(conflictMessage, vbAbortRetryIgnore)
        
        If ri = vbAbort Then '[中止]ボタンが押された
            Exit Sub
        
        ElseIf ri = vbRetry Then '[再試行]ボタンが押された
            ignoreVacant = False '空文字も上書きする
            
        Else '[無視]ボタンが押された
            ignoreVacant = True '空文字は無視する
            
        End If
    
    End If
    
    'ペーストループ
    For lineCounter = 0 To maxNumOfLines '行ループ
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        numOfColumns = UBoundSafe(toPasteStrings)
        
        If numOfColumns < 0 Then '空配列だった場合
            toPasteStrings = Array("") '空文字1つの配列を定義する
            numOfColumns = 0

        End If
        
        For columnCounter = 0 To numOfColumns '列ループ
            
            '結合セルの場合は、結合セルの左上にのみペーストする
            If (Selection(1).Offset(lineCounter, columnCounter).Address = Selection(1).Offset(lineCounter, columnCounter).MergeArea.Cells(1, 1).Address) Then
                
                '空文字の場合は、無視設定されていなかれば貼り付ける
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


'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>
 
