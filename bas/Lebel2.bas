Attribute VB_Name = "Lebel2"
'Excel操作　＋　条件分岐
Sub Lebel2()
    '点数を入れる変数
    Dim score As Integer
    
    'テスト用の点数
    score = 85
    
    'B2セルに点数を表示
    Range("B2").Value = score
    
    '判定処理
    If score >= 90 Then
        MsgBox "すごい！A評価です！"
    ElseIf score >= 70 Then
        MsgBox "いい感じ！B評価です！"
    Else
        MsgBox "次は頑張ろう！C評価です！"
    End If
    
    '終了メッセージ
    MsgBox "Level2 完了！"
    
End Sub
