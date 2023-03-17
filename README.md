# VBA_challenge
UC Berkeley Data Analytics Bootcamp Module 2 Challenge
--
This is the solution:
-

Sub Module_2_JPF_Solved_Final_Version()

'  Enable script to run on every worksheet at once
For Each ws In Worksheets

    Dim TotalRecords As Double
    Dim Counter As Integer
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Dim GPIValue As Double
    Dim GPDValue As Double
    Dim GTValue As Double
    Dim GPITicker As String
    Dim GPDTicker As String
    Dim GTVTicker As String
    Dim TotalRecords2 As Double
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    TotalRecords = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Counter = 1
    
    OpeningPrice = ws.Range("C2").Value
    
    Volume = 0
    
    ws.Range("I2").Value = ws.Range("A2").Value
    
    For i = 2 To TotalRecords
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' The ticker symbol
            ws.Cells(2 + Counter, 9).Value = ws.Cells(i + 1, 1).Value
                
            ' Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    
            ClosingPrice = ws.Cells(i, 6).Value
                
            YearlyChange = ClosingPrice - OpeningPrice
            
            ws.Cells(1 + Counter, 10).Value = YearlyChange
            
            ' Conditional formatting (Yearly Change) highlight positive change in green and negative change in red
                If YearlyChange > 0 Then
                
                    ws.Cells(1 + Counter, 10).Interior.Color = vbGreen
                
                ElseIf YearlyChange < 0 Then
                
                    ws.Cells(1 + Counter, 10).Interior.Color = vbRed
                
                End If
            
            ' The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

            PercentChange = YearlyChange / OpeningPrice
            
            ws.Cells(1 + Counter, 11).Value = FormatPercent(PercentChange)
            
            ' Conditional formatting (Percent Change) highlight positive change in green and negative change in red
                If PercentChange > 0 Then
                
                    ws.Cells(1 + Counter, 11).Interior.Color = vbGreen
                
                ElseIf PercentChange < 0 Then
                
                    ws.Cells(1 + Counter, 11).Interior.Color = vbRed
                
                End If
            
            ' The total stock volume of the stock.
            Volume = Volume + ws.Cells(i, 7).Value
            
            ws.Cells(1 + Counter, 12).Value = Volume
            
            Counter = Counter + 1
            
            OpeningPrice = ws.Cells(i + 1, 3).Value
            
            Volume = 0
        
        Else
        
            Volume = Volume + ws.Cells(i, 7).Value
            
        End If
    
    Next i
    
    ws.Columns("I:L").AutoFit
    
    ' Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    GPIValue = ws.Range("K2").Value
    GPDValue = ws.Range("K2").Value
    GTVValue = ws.Range("L2").Value
    
    TotalRecords2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To TotalRecords2
    
        If ws.Cells(i, 11).Value > GPIValue Then
        
            GPIValue = ws.Cells(i, 11).Value
            GPITicker = ws.Cells(i, 9).Value
            
        End If
        
        If ws.Cells(i, 11).Value < GPDValue Then
        
            GPDValue = ws.Cells(i, 11).Value
            GPDTicker = ws.Cells(i, 9).Value
        
        End If
        
        If ws.Cells(i, 12).Value > GTVValue Then
        
            GTVValue = ws.Cells(i, 12).Value
            GTVTicker = ws.Cells(i, 9).Value
        
        End If
    
    Next i
    
    ' Greatest % increase
    ws.Range("P2").Value = GPITicker
    ws.Range("Q2").Value = FormatPercent(GPIValue)
    
    ' Greatest % decrease
    ws.Range("P3").Value = GPDTicker
    ws.Range("Q3").Value = FormatPercent(GPDValue)
    
    ' Greatest total volume
    ws.Range("P4").Value = GTVTicker
    ws.Range("Q4").Value = GTVValue
    
    ws.Columns("O:Q").AutoFit

Next ws

End Sub
