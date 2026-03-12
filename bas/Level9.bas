Attribute VB_Name = "Level9"
Sub Level9_åüçı()

    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    Dim found As Boolean
    
    keyword = Range("D2").Value
    found = False
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
    
        If Cells(i, 1).Value = keyword Then
        
            MsgBox i & "çsñ⁄Ç…å©Ç¬Ç©ÇËÇ‹ÇµÇΩ"
            found = True
            Exit For
        
        End If
    
    Next i
    
    If found = False Then
        MsgBox "å©Ç¬Ç©ÇËÇ‹ÇπÇÒÇ≈ÇµÇΩ"
    End If
    
End Sub
