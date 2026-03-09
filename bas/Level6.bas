Attribute VB_Name = "Level6"
Sub Level6_çáåv()
    
    Dim i As Long
    Dim total As Long
    
    total = 0
    
    For i = 2 To 5
        
        total = total + Cells(i, 1).Value
        
    Next i
    
    Cells(6, 1).Value = total

End Sub
