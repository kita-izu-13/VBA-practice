Attribute VB_Name = "Level11"
Sub Level11_èdï°É`ÉFÉbÉN()

    Dim lastRow As Long
    Dim i As Long
    Dim song As String
    Dim artist As String
    
    song = Range("D2").Value
    artist = Range("D3").Value
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
    
        If Cells(i, 1).Value = song And Cells(i, 2).Value = artist Then
        
            MsgBox "ä˘Ç…ìoò^Ç≥ÇÍÇƒÇ¢Ç‹Ç∑"
            Exit Sub
            
        End If
        
    Next i
    
    'ìoò^
    Cells(lastRow + 1, 1).Value = song
    Cells(lastRow + 1, 2).Value = artist
    
    MsgBox "ìoò^ÇµÇ‹ÇµÇΩ"

End Sub
