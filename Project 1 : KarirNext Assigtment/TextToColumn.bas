Attribute VB_Name = "Module1"
Sub TexttoCol_Dynamic()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    
    Range(Selection, Selection.End(xlDown)).TextToColumns _
        Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, _
        FieldInfo:=Array(Array(1, 1), Array(2, 4), Array(3, 1), Array(4, 1), Array(5, 1), _
                         Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), _
                         Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), _
                         Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1)), _
        TrailingMinusNumbers:=True
        
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ws.Range("U1").Value = "Bulan"
    
    For i = 2 To lastRow
        If ws.Cells(i, "B").Value <> "" Then
            ws.Cells(i, "U").FormulaR1C1 = "=TEXT(RC[-19],""MMM"")"
        Else
            ws.Cells(i, "U").Value = ""
        End If
    Next i
    
    MsgBox "Selesai! Data Sudah Berhasil Di Split.", vbInformation, "Sukses"
End Sub
