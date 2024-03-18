VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Year2018Stock_data():

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

Open_price = Cells(2, 3).Value

' Set an initial variable for holding the Close_price

Dim Close_price As Double

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

 ' Loop through all Tickers

For i = 2 To Lastrow

 ' Check if we are still within the same Ticker, if it is not...

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

 ' Set the Brand name and Close price
Ticker = Cells(i, 1).Value
Close_price = Cells(i, 6).Value

'Calculate for the Yearly_change and Percent_change

Yearly_change = (Close_price - Open_price)
Percent_change = (Yearly_change / Open_price)

' Add to the Total Stock Volume
Total_stockVol = Total_stockVol + Cells(i, 7).Value

 ' Print the Ticker in the Summary Table
Range("K" & Summary_Table_Row).Value = Ticker

 ' Print the Total Stock Volume to the Summary Table
Range("N" & Summary_Table_Row).Value = Total_stockVol

 ' Print the Yearly change to the Summary Table and format to 2 decimal places
Range("L" & Summary_Table_Row).Value = Yearly_change
Range("L" & Summary_Table_Row).NumberFormat = "0.00"

' Print the Percent_change to the Summary Table and format to 2 decimal places

Range("M" & Summary_Table_Row).Value = Percent_change
Range("M" & Summary_Table_Row).NumberFormat = "0.00%"

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

' Reset the Open Price and Total Stock Volume
Total_stockVol = 0
Open_price = Cells(i + 1, 3).Value

' If the cell immediately following a row is the same Ticker...
Else

' Add to the Ticker Total
Total_stockVol = Total_stockVol + Cells(i, 7).Value

End If

Next i

' Obtaining the rowcount for Yearly_change in summary Table

Yearly_change_Lastrow = Cells(Rows.Count, 12).End(xlUp).Row

' Loop through the Yearly_change row

For i = 2 To Yearly_change_Lastrow

'If the value in this row is > than 0 then make the interior red

If Cells(i, 12).Value < 0 Then
    Cells(i, 12).Interior.ColorIndex = 3

    
    Else
    
    Cells(i, 12).Interior.ColorIndex = 4

End If

Next i


End Sub


