Attribute VB_Name = "Level4"
Sub Level4()

    Dim lastRow As Long
    Dim song As String
    Dim artist As String
    
    '入力値を取得
    song = Range("D2").Value
    artist = Range("D3").Value
    
    '空チェック
    If song = "" Or artist = "" Then
        MsgBox "曲名と歌手名を入力してください"
        Exit Sub
    End If
    
    '最終行取得
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    '追加
    Cells(lastRow, 1).Value = song
    Cells(lastRow, 2).Value = artist
    
    '入力欄クリア
    Range("D2").ClearContents
    Range("D3").ClearContents
    
    MsgBox "追加しました！"

End Sub

