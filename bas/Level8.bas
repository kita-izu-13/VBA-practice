Attribute VB_Name = "Level8"
Sub Level8_複数抽選()

    Dim wsList As Worksheet
    Dim wsHist As Worksheet
    Dim lastRowList As Long
    Dim lastRowHist As Long
    Dim pick As Long
    Dim result As String
    
    Set wsList = Sheets("8")
    Set wsHist = Sheets("8-1")
    
    'リスト最後の行
    lastRowList = wsList.Cells(wsList.Rows.Count, 1).End(xlUp).Row
    
    '空なら終了
    If lastRowList < 1 Then
        MsgBox "リストが空です"
        Exit Sub
    End If
    
    'ランダム抽選
    Randomize
    pick = Int(Rnd * lastRopwList) + 1
    
    '選ばれた名前
    result = wsList.Cells(pick, 1).Value
    
    '履歴シートに記録
    lastRowHist = wsHist.Cells(wsHist.Rows.Count, 1).End(xlUp).Row + 1
    wsHist.Cells(lastRowHist, 1).Value = Date
    wsHist.Cells(lastRowHist, 2).Value = result
    
    'リストから削除（次回抽選に入らない）
    wsList.Rows(pick).Delete
    
    '結果を表示
    MsgBox "今日の当番 → " & result

End Sub
