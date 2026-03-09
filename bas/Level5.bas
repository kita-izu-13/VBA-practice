Attribute VB_Name = "Level5"
Sub Level5()

    Dim song As String
    Dim artist As String
    Dim lastRow As Long
    Dim i As Long
    
    '入力取得
    song = Range("D2").Value
    artist = Range("D3").Value
    
    '空チェック
    If song = "" Or artist = "" Then
        MsgBox "曲名と歌手名を入力してください"
        Exit Sub
    End If
    
    '最終行取得
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '重複チェック
    For i = 2 To lastRow
        If Cells(i, 1).Value = song And Cells(i, 2).Value = artist Then
            MsgBox "この曲はすでに登録されています"
            Exit Sub
        End If
    Next i
    
    '追加
    Cells(lastRow + 1, 1).Value = song
    Cells(lastRow + 1, 2).Value = artist
    
    '入力欄クリア
    Range("D2").ClearContents
    Range("D3").ClearContents
    
    MsgBox "追加しました！"

End Sub
