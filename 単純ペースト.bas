Attribute VB_Name = "単純ペースト"

Sub 単純ペースト()

    Dim CB As New DataObject
    
    Dim delimiter As String: delimiter = vbTab 'セル間区切り文字
    Dim NLCharacter As String: NLCharacter = vbCrLf '改行文字
    
    Dim maxNumOfLines As Long
    Dim maxNumOfColumns As Long
    
    Dim cautionMessage As String: cautionMessage = "このSubプロシージャは、" & vbLf & _
                                                   "現在の選択範囲に対して値の書き込みを行います。" & vbLf & vbLf & _
                                                   "実行しますか?"
    
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
    
    CB.GetFromClipboard 'クリップボードの内容を取得
    
    If Not (CB.GetFormat(1)) Then 'テキスト形式でない場合
        retVal = MsgBox("クリップボードにテキストがありません", vbExclamation)
        Exit Sub '終了
        
    End If
    
    linesOfToPasteText = Split(CB.GetText, NLCharacter) '行区切りの文字列配列を取得
    
    maxNumOfLines = UBound(linesOfToPasteText)
    maxNumOfColumns = 0
    
    '上書き確認用最大列数の検査
    For lineCounter = 0 To maxNumOfLines '行ループ for 最大column数検査
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        numOfColumns = UBound(toPasteStrings)
        
        If (maxNumOfColumns < numOfColumns) Then
        
            maxNumOfColumns = numOfColumns '最大列数の保存
            
        End If
    
    Next lineCounter
    
    '書込み先セル範囲の選択
    Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row + maxNumOfLines, Selection.Column + maxNumOfColumns)).Select
    
    '上書き確認
    If WorksheetFunction.CountA(Selection) > 0 Then
        yn = MsgBox("ペースト先のセルに値が入っています" & vbLf & vbLf & _
                    "上書きしますか？", _
                    vbOKCancel)
        
        If yn = vbCancel Then
            Exit Sub
        End If
    
    End If
    
    'ペーストループ
    For lineCounter = 0 To maxNumOfLines '行ループ
        
        toPasteStrings = Split(linesOfToPasteText(lineCounter), delimiter)
        
        numOfColumns = UBound(toPasteStrings)
        
        For columnCounter = 0 To numOfColumns '列ループ
            
            '結合セルの場合は、結合セルの左上にのみペーストする
            If (Selection(1).Offset(lineCounter, columnCounter).Address = Selection(1).Offset(lineCounter, columnCounter).MergeArea.Cells(1, 1).Address) Then
                
                Selection(1).Offset(lineCounter, columnCounter).Value = toPasteStrings(columnCounter)
                
            End If
            
        Next columnCounter
    
    Next lineCounter
    
End Sub

