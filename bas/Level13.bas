Attribute VB_Name = "Level13"
Sub Level13_条件コピー()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim pasteRow As Long
    Dim i As Long
    
    Set ws1 = Sheets("13")
    Set ws2 = Sheets("13-1")
    
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    pasteRow = 1
    
    For i = 2 To lastRow
    
        If ws1.Cells(i, 2).Value = "米津玄師" Then
        
            ws2.Cells(pasteRow, 1).Value = ws1.Cells(i, 1).Value
            ws2.Cells(pasteRow, 2).Value = ws1.Cells(i, 2).Value
            
            pasteRow = pasteRow + 1
    
        End If
    
    Next i
    
    MsgBox "抽出完了"

End Sub
