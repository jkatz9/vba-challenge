Sub grubbySquirrel()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim changeIndices() As Long ' Array to hold the row indices where the ticker symbol changes
    Dim changeCount As Long ' Number of times the ticker symbol changes (the size of the changeIndices array)

    For Each ws In ThisWorkbook.Worksheets ' Loop through each worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Get the last row of the worksheet for looping purposes
        changeCount = 0 ' Initialize the change count

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        For i = 2 To lastRow ' Loop through rows
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then ' If the ticker symbol changes
                changeCount = changeCount + 1 ' Increment the change count
                ReDim Preserve changeIndices(1 To changeCount) ' Resize the array to hold the row indices where the ticker symbol changes
                changeIndices(changeCount) = i ' Record the row index where the ticker symbol changes
            End If
        Next i

        For i = 1 To changeCount - 1
            ws.Cells(i + 1, 9) = ws.Cells(changeIndices(i), 1) ' Set ticker symbol
            ws.Cells(i + 1, 10) = ws.Cells(changeIndices(i + 1) - 1, 6) - ws.Cells(changeIndices(i), 3) ' Closing price minus initial price
            If ws.Cells(i + 1, 10).Value < 0 Then
                ws.Cells(i + 1, 10).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                ws.Cells(i + 1, 10).Interior.Color = RGB(0, 255, 0) ' Green
            End If
            ws.Cells(i + 1, 11) = ws.Cells(i + 1, 10) / ws.Cells(changeIndices(i), 3) ' Percent change
            ws.Cells(i + 1, 12) = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(changeIndices(i), 7), ws.Cells(changeIndices(i + 1) - 1, 7))) ' Sum the volumes for the quarter
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestTotalVolume As Double
        greatestIncrease = WorksheetFunction.Max(ws.Range("K:K"))
        greatestDecrease = WorksheetFunction.Min(ws.Range("K:K"))
        greatestTotalVolume = WorksheetFunction.Max(ws.Range("L:L"))

        For i = 1 To changeCount
            If ws.Cells(i + 1, 11) = greatestIncrease Then ' If the percent change is the greatest increase
                ws.Cells(2, 16) = ws.Cells(i + 1, 9) ' Set ticker symbol
                ws.Cells(2, 17) = ws.Cells(i + 1, 11) ' Set value
                ws.Cells(2, 17).NumberFormat = "0.00%" ' Format as percentage
            End If
            If ws.Cells(i + 1, 11) = greatestDecrease Then ' If the percent change is the greatest decrease
                ws.Cells(3, 16) = ws.Cells(i + 1, 9) ' Set ticker symbol
                ws.Cells(3, 17) = ws.Cells(i + 1, 11) ' Set value
                ws.Cells(3, 17).NumberFormat = "0.00%" ' Format as percentage
            End If
            If ws.Cells(i + 1, 12) = greatestTotalVolume Then ' If the total volume is the greatest
                ws.Cells(4, 16) = ws.Cells(i + 1, 9) ' Set ticker symbol
                ws.Cells(4, 17) = ws.Cells(i + 1, 12) ' Set value
            End If
        Next i

    Next ws ' On to the next worksheet
End Sub
