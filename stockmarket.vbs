
Sub stockmarket()

    'loop through different worksheets
    Dim ws As Worksheet

        For Each ws In Worksheets
        
            'column names for assignment result
            ws.Cells(1, 9).Value = "Ticker Name"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

            Dim Ticker_Name As String
            Dim Year_Change As Double
            Dim Percent_Change As Double
            Dim Total_Stock As Double

            Total_Stock = 0
            Percent_Change = 0

            Dim Year_Open As Double
            Dim Year_Close As Double

            Year_Open = 0
            Year_Close = 0
            ' after the header, start from row 2
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
            'total rows for looping
                Dim LastRow As Long
                LastRow = Cells(Rows.Count, 1).End(xlUp).Row

                For i = 2 To LastRow
                    'year open based on previous ticker change 
                    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                        Year_Open = ws.Cells(i, 3).Value
                    End If

                    ' if the ticker name is changing, total stock
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        Ticker_Name = ws.Cells(i, 1).Value
                            'total up the stock volume
                        Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                            'column I for ticker colum L for total stock 
                        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                        ws.Range("L" & Summary_Table_Row).Value = Total_Stock

                        Year_Close = ws.Cells(i, 6).Value
                        Year_Change = Year_Close - Year_Open
                        ws.Range("J" & Summary_Table_Row).Value = Year_Change

                            'conditional formatting for positive/negative
                        If Year_Change >= 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                            'if yearly open and yearly close is 0 then percent change also is 0 
                        If Year_Open = 0 And Year_Close = 0 Then
                        Percent_Change = 0
                            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                            ws.Range("K" & Summary_Table_Row).Value = "0.00%"
                            'any increase from 0 is going to display as infinity - looking only at increase by dollar amount 
                        ElseIf Year_Open = 0 Then
                            Dim percent_change_zeroincrease As String
                            percent_change_zeroincrease = "ZeroIncrease"
                            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        Else
                            'percent change is simply yearly change from open to close divide by yearly open 
                            Percent_Change = Year_Change / Year_Open
                            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        End If
                            'looking at the next row in the summary table 
                        Summary_Table_Row = Summary_Table_Row + 1
                            'resetting to 0 for count 
                        Total_Stock = 0
                        Year_Open = 0
                        Year_Close = 0
                        Year_Change = 0
                        Percent_Change = 0

                        End If
    
                Next i
    Next ws
End Sub