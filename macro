Sub CopyDataToDSBasedOnHighlight()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Specify the worksheet where your data is located
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' Change "Sheet1" to your sheet's name

    ' Find the last row with data in column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

    ' Loop through each row in column I
    For i = 1 To lastRow
        ' Check if the cell in column I is highlighted
        If ws.Cells(i, "I").Interior.ColorIndex <> xlNone Then
            ' If highlighted, print "yes" into column DS
            ws.Cells(i, "DS").Value = "yes"
        Else
            ' If not highlighted, print "no" into column DS
            ws.Cells(i, "DS").Value = "no"
        End If
    Next i
End Sub
