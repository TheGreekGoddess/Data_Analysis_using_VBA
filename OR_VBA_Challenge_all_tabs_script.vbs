Sub StockMarketAnalysisAllTabs():
        For Each ws In Worksheets
        ws.Activate
' Set the initial variable for holding the different data sets in the Summary table
            Dim Ticker_Name As String
            Dim Opening_Price As Double
                Opening_Price = 0
            Dim Closing_Price As Double
                Closing_Price = 0
            Dim Yearly_Change As Double
                Yearly_Change = 0
            Dim Percent_Change As Double
                Percent_Change = 0
            Dim Total_Stock_Volume As Double
                Total_Stock_Volume = 0
            Dim LastRow As Double
            
' List unique tickers in the Summary table
            Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
            
' Set the summary table of Ticker Names
            Range("I1").Value = "Ticker Name"
            Range("J1").Value = "Yearly Change $"
            Range("K1").Value = "Percent Change %"
            Range("L1").Value = "Total Stock Volume"
        
' Set the first and last row
            FirstRow = 2
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow
                    
' Count all similar tickers and move to next ticker
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                        
' Publish and format the Ticker Name in the Summary Table
                    Ticker_Name = Cells(i, 1).Value
                    Range("I" & Summary_Table_Row).Value = Ticker_Name

' Publish and format the Yearly Change to the Summary Table
                    Opening_Price = Cells(FirstRow, 3).Value
                    Closing_Price = Cells(i, 6).Value
                    Yearly_Change = (Closing_Price - Opening_Price)
                    Range("J" & Summary_Table_Row).Value = Yearly_Change
                    Range("J" & Summary_Table_Row).HorizontalAlignment = xlRight
                    Range("J" & Summary_Table_Row).NumberFormat = "$#.##"
                    If Range("J" & Summary_Table_Row).Value < 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                        
' Publish and format the Percent Change to the Summary Table
                    Percent_Change = (Yearly_Change / Opening_Price)
                    Range("K" & Summary_Table_Row).Value = Percent_Change
                    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    If Range("K" & Summary_Table_Row).Value < 0 Then
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                        
' Publish and format the Total Stock Volume to the Summary Table
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                    Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                    Range("L" & Summary_Table_Row).NumberFormat = "#,##"
                    Range("L" & Summary_Table_Row).HorizontalAlignment = xlRight

' Add the next Ticker Name to the summary table
                    Summary_Table_Row = Summary_Table_Row + 1
        
' Reset the Total Stock Volume for next Ticker Name
                    Yearly_Change = 0
                    Percent_Change = 0
                    Total_Stock_Volume = 0
                    StartingRow = i + 1
                Else
                    
' Add up the Total Stock Volume for same Ticker Name
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                End If
            Next i
                                    
' For +10 Mastery points set the initial variables for greatest calculations table
            Dim Greatest_Increase As Double
            Dim Greatest_Decrease As Double
            Dim Greatest_Total_Volume As Double
            Dim Greatest_Increase_Ticker_Name As String
            Dim Greatest_Decrease_Ticker_Name As String
            Dim Greatest_Total_Volume_Ticker_Name As String
            Dim LastRowOfSummary As Double

' Set the calculations for Greatest Increase & Decrease in Percent Change and Total Volume
            Greatest_Increase = Application.WorksheetFunction.Max(Range("K:K"))
            Greatest_Decrease = Application.WorksheetFunction.Min(Range("K:K"))
            Greatest_Total_Volume = Application.WorksheetFunction.Max(Range("L:L"))
            
' Set the loop to calculate the Greatest values
            LastRowOfSummary = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To LastRowOfSummary
                If ws.Cells(i, 11) = Greatest_Increase Then
                    Greatest_Increase_Ticker_Name = ws.Cells(i, 11).Offset(0, -2).Value
                End If
                    
                If ws.Cells(i, 11) = Greatest_Decrease Then
                    Greatest_Decrease_Ticker_Name = ws.Cells(i, 11).Offset(0, -2).Value
                End If
                
                If ws.Cells(i, 12) = Greatest_Total_Volume Then
                    Greatest_Total_Volume_Ticker_Name = ws.Cells(i, 12).Offset(0, -3).Value
                End If
            Next i
            
' Set, publish and format the "Greatest" table of Ticker Names
            Range("O1").Value = "Analysis of the Greatest:"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("O:O").HorizontalAlignment = xlLeft
            Range("P1").Value = "Ticker Name"
            Range("Q1").Value = "Value"
            
            Range("P2").Value = Greatest_Increase_Ticker_Name
            Range("Q2").Value = Greatest_Increase
            Range("Q2").NumberFormat = "0.00%"

            Range("P3").Value = Greatest_Decrease_Ticker_Name
            Range("Q3").Value = Greatest_Decrease
            Range("Q3").NumberFormat = "0.00%"
            
            Range("P4").Value = Greatest_Total_Volume_Ticker_Name
            Range("Q4").Value = Greatest_Total_Volume
            Range("Q4").NumberFormat = "#,##"
            
            Range("A1:Q1").Font.FontStyle = "Bold"
        Next ws
End Sub