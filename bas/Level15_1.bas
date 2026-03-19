Attribute VB_Name = "Level15_1"
Sub Level15_追加()

    Dim lastRow As Long
    Dim i As Long
    Dim name As String
    Dim memo As String
    
    name = Range("E1").Value
    memo = Range("E2").Value
    
    '空チェック
    If name = "" Then
        MsgBox "名前を入力してください"
        Exit Sub
    End If
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '重複チェック
    For i = 2 To lastRow
        If Cells(i, 1).Value = name Then
            MsgBox "既に登録されています"
            Exit Sub
        End If
    Next i
    
    '追加
    Cells(lastRow + 1, 1).Value = name
    Cells(lastRow + 1, 2).Value = memo
    
    MsgBox "登録しました"

End Sub
