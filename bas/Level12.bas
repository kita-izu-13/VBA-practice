Attribute VB_Name = "Level12"
Sub Level12_別シートコピー()

    Dim lastRow As Long
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Set ws1 = Sheets("12")
    Set ws2 = Sheets("12-1")
    
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    
    ws1.Range("A1:B" & lastRow).Copy
    ws2.Range("A1").PasteSpecial xlPasteValues
    
    MsgBox "コピーしました"
    

End Sub
