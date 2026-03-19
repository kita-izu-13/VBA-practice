Attribute VB_Name = "Level15_2"
Sub Level15_ŒŸõ()

    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    Dim result As String
    
    keyword = Range("E4").Value
    result = ""
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
    
        If InStr(Cells(i, 1).Value, keyword) > 0 Then
        
            result = result & Cells(i, 1).Value & " / " & Cells(i, 2).Value & vbCrLf
        
        End If
    
    Next i
    
    If result = "" Then
        MsgBox "Œ©‚Â‚©‚è‚Ü‚¹‚ñ‚Å‚µ‚½"
    Else
        MsgBox result
    End If
    
End Sub
