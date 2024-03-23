Attribute VB_Name = "Year2018To2022"
Sub Year2018To2020Year_Stock():

For Each ws In Worksheets

' Set an initial variable for holding the Ticker, Yearly_change, Percent_change, Total_stockVol

Dim Ticker As String
Dim Yearly_change As Double
Dim Percent_change As Double
Dim Total_stockVol As Double
Total_stockVol = 0

' Keep track of the location for each ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Set an initial variable for holding the Open_price

Dim Open_price As Double

'-----Open price at the beginning of the year

Open_price = ws.Cells(2, 3).Value

' Set an initial variable for holding the Close_price

Dim Close_price As Double

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 11).Value = "Ticker"

ws.Cells(1, 12).Value = "Yearly Change"

ws.Cells(1, 13).Value = "Percent Change"

ws.Cells(1, 14).Value = "Total Stock Volume"


 ' Loop through all Tickers

For i = 2 To Lastrow

 ' Check if we are still within the same Ticker, if it is not...

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

 ' Set the Ticker and Close price
Ticker = ws.Cells(i, 1).Value
Close_price = ws.Cells(i, 6).Value

'Calculate for the Yearly_change and Percent_change

Yearly_change = (Close_price - Open_price)
Percent_change = (Yearly_change / Open_price)

' Add to the Total Stock Volume
Total_stockVol = Total_stockVol + ws.Cells(i, 7).Value

 ' Print the Ticker in the Summary Table
ws.Range("K" & Summary_Table_Row).Value = Ticker

 ' Print the Total Stock Volume to the Summary Table
ws.Range("N" & Summary_Table_Row).Value = Total_stockVol

 ' Print the Yearly change to the Summary Table and format to 2 decimal places
ws.Range("L" & Summary_Table_Row).Value = Yearly_change
ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00"

' Print the Percent_change to the Summary Table and format to 2 decimal places

ws.Range("M" & Summary_Table_Row).Value = Percent_change
ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Reset the Open Price and Total Stock Volume
Total_stockVol = 0
Open_price = ws.Cells(i + 1, 3).Value

' If the cell immediately following a row is the same Ticker...
Else

' Add to the Ticker Total
Total_stockVol = Total_stockVol + ws.Cells(i, 7).Value

End If

Next i

' Obtaining the rowcount for Yearly_change in summary Table

Yearly_change_Lastrow = ws.Cells(Rows.Count, 12).End(xlUp).Row

' Loop through the Yearly_change row

For i = 2 To Yearly_change_Lastrow

'If the value in this row is > than 0 then make the interior red

If ws.Cells(i, 12).Value < 0 Then
    ws.Cells(i, 12).Interior.ColorIndex = 3

    
    Else
    
    ws.Cells(i, 12).Interior.ColorIndex = 4

End If

Next i

'Label the cell rows for the asks: Greatest % Increase, Greatest % Decrease, Greatest Total Volume, Ticker, and Value
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"

'Obtaining the rowcount for Percent_change in summary Table

Percent_change_Lastrow = ws.Cells(Rows.Count, 13).End(xlUp).Row

'Obtaining the rowcount for Total_StockVol in summary Table
Total_StockVol_Lastrow = ws.Cells(Rows.Count, 14).End(xlUp).Row

' Loop through all Percent_Change row in summary Table

For i = 2 To Percent_change_Lastrow

' Calculate and format the Greatest % Increase in Percent_Change row

If ws.Cells(i, 13).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & Percent_change_Lastrow)) Then

ws.Range("Q2").Value = ws.Cells(i, 11).Value
ws.Range("R2").Value = ws.Cells(i, 13).Value
ws.Range("R2").NumberFormat = "0.00%"

' Calculate and format the Greatest % decrease in Percent_Change row using Elseif conditional function

ElseIf ws.Cells(i, 13).Value = Application.WorksheetFunction.Min(ws.Range("M2:M" & Percent_change_Lastrow)) Then


ws.Range("Q3").Value = ws.Cells(i, 11).Value
ws.Range("R3").Value = ws.Cells(i, 13).Value
ws.Range("R3").NumberFormat = "0.00%"


' Calculate and format the Greatest Total_StockVol in Total_StockVol row
ElseIf ws.Cells(i, 14).Value = Application.WorksheetFunction.Max(ws.Range("N2:N" & Total_StockVol_Lastrow)) Then
ws.Range("Q4").Value = ws.Cells(i, 11).Value
ws.Range("R4").Value = ws.Cells(i, 14).Value

End If
Next i

Next ws


End Sub

