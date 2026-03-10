Attribute VB_Name = "Level7"
Sub Level7_ランダム抽選()

    Dim lastRow As Long
    Dim pick As Long
    Dim result As String
    
    'リストの最後の行を取得
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'ランダムに1～lastRowの数を選ぶ
    Randomize   '乱数の初期化
    pick = Int(Rnd * (lastRow - 1 + 1)) + 1 '1～lastRow
    
    '選んだ名前を取得
    result = Cells(pick, 1).Value
    
    '結果を表示
    MsgBox "今日の当番→" & result

End Sub
