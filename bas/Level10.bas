Attribute VB_Name = "Level10"
Sub Level10_•”•ªˆê’vŒŸõ()

    Dim lastRow As Long
    Dim i As Long
    Dim keyword As String
    Dim result As String
    Dim found As Boolean
    
    keyword = Range("D2").Value
    found = False
    result = ""
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastRow

        
        '•”•ªˆê’vŒŸõ
        If InStr(Cells(i, 1).Value, keyword) > 0 Then
            result = result & i & "s–Ú ¨ " & Cells(i, 1).Value & vbCrLf
            found = True
        End If
    
    Next i
    
    If found Then
        MsgBox "ƒqƒbƒg‚µ‚½‚à‚Ì:" & vbCrLf & result
    
    Else
        MsgBox "Œ©‚Â‚©‚è‚Ü‚¹‚ñ‚Å‚µ‚½"
    End If

End Sub
