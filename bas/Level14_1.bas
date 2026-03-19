Attribute VB_Name = "Level14_1"
Sub Level14_ANDíäèo()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim pasteRow As Long
    Dim i As Long
    
    Set ws1 = Sheets("14")
    Set ws2 = Sheets("14-1")
    
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    pasteRow = 1
    
    For i = 2 To lastRow
    
        If ws1.Cells(i, 2).Value = "ïƒí√å∫ét" _
            And InStr(ws1.Cells(i, 1).Value, "Ép") > 0 Then
        
            ws2.Cells(pasteRow, 1).Value = ws1.Cells(i, 1).Value
            ws2.Cells(pasteRow, 2).Value = ws1.Cells(i, 2).Value
            
            pasteRow = pasteRow + 1
    
        End If
    
    Next i
    
    MsgBox "íäèoäÆóπ"

End Sub
