Attribute VB_Name = "Module1"
Sub VBA_Stock_Analysis()

    'Create the variables.
    Dim i As Long
    Dim last_row As Long
    
    'Create the variables for column and summary row.
    Dim column As Integer
    Dim summary_row As Integer
    
    'Create the variable for ticker name.
    Dim ticker_name As String
    
    'Create the variables for Yearly Change, Opening Price, Closing Price, Percent Change and Total Stock Volume.
    Dim Yearly_Change As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    
Dim ws As Worksheet
For Each ws In Worksheets
ws.Select
    
    'Create column name for Ticker, Yearly Change, Percent Change and Total Stock Volume.
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Create Summmary Table worksheet for Greatest % Increase, Greastest % Decrease and Greatest Total Volume.
    ws.Range("O1").Value = "Column Name"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Create Summary Table worksheet for Ticker and Value.
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Using percentage formatting.
    ActiveSheet.Range("K2").Select
    ActiveSheet.Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00%"
    
column = 1
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
Opening_Price = ws.Cells(2, 3).Value
    
'Starting the counter.
summary_row = 2
total_vol = 0
    
    'Looping through rows in the 1st column of wholesheet.
    For i = 2 To last_row
    
    total_vol = total_vol + ws.Cells(i, 7).Value
    
    'Declaring if column 1 is different from another column.
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
    
    'Create ticker_name in column 9.
    ticker_name = ws.Cells(i, column).Value
    ws.Cells(summary_row, 9).Value = ticker_name
    
    'Calculating yearly change & percent_change
    Closing_Price = ws.Cells(i, 6).Value
    Yearly_Change = Closing_Price - Opening_Price
    
    'Set Yearly Change in column 10.
    ws.Cells(summary_row, 10).Value = Yearly_Change
    
    If (Opening_Price > 0) Then
    Percent_Change = (Yearly_Change / Opening_Price)
    Else
    Percent_Change = 0
    
    End If
    
    'Set Percent Change in column 11.
    ws.Cells(summary_row, 11).Value = Percent_Change
    
    'set the total_vol in column 12
    ws.Cells(summary_row, 12).Value = total_vol

    'Conditionally format column with colors base on positive or negative.
    If Yearly_Change > 0 Then
    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
    
    End If
    
    Opening_Price = ws.Cells(i + 1, 3).Value
    total_vol = 0
    summary_row = summary_row + 1
        
    End If
    
    Next i
    
    'Calculating Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    'Start Looping for Final Results
    For i = 2 To LastRow
    
    'Greatest % Increase
    If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
    ws.Range("Q2").Value = ws.Range("K" & i).Value
    ws.Range("P2").Value = ws.Range("I" & i).Value
    
    End If
    
    'Greatest % Decrease
    If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Range("K" & i).Value
    ws.Range("P3").Value = ws.Range("I" & i).Value
    
    'Greatest Total Volume
    If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
    ws.Range("Q4").Value = ws.Range("L" & i).Value
    ws.Range("P4").Value = ws.Range("I" & i).Value
    
    End If
    
    'Using percentage formatting
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

    End If

    Next i

Next ws
    
End Sub
