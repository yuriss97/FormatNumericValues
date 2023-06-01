Function FormatValuesWithDotsAndCommas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) 
    
    With ws.Cells
        .Font.Name = "Arial Narrow"
        .Font.Size = 10
    End With

    ws.UsedRange.NumberFormat = "#,##0.00"
End Function
