Sub yearstock():

    Dim tickersymbol As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalstockvolume As Double
    Dim openning As Double
    Dim closing As Double
    Dim lastrow As Long
    Dim Summaryrow As Integer
    Dim i As Long
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ' Reset variables for each worksheet
        totalstockvolume = 0
        openning = 0
        closing = 0
        Summaryrow = 2
        
        ' Set headers
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Volume"
        
        ' Find the last row with data in column A
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each row
        For i = 2 To lastrow
        
            ' Check if it's the last row for the current ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                ' Calculate yearly change and percentage change
                yearlychange = ws.Cells(i, 3).Value - ws.Cells(Start, 3).Value
                percentchange = yearlychange / ws.Cells(Start, 3).Value
                
                ' Update total volume
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                
                ' Write values to the summary row
                ws.Cells(Summaryrow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Summaryrow, 10).Value = yearlychange
                ws.Cells(Summaryrow, 11).Value = percentchange
                ws.Cells(Summaryrow, 12).Value = totalstockvolume
                
                ' Reset variables for the next ticker
                totalstockvolume = 0
                Start = i + 1
                Summaryrow = Summaryrow + 1
            Else
                ' Increment total volume
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            End If
        
        Next i
        
    Next ws

End Sub
