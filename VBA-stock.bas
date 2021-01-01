Attribute VB_Name = "Module1"
Sub stock_data():


  ' Set variables
Dim Ticker As String
Dim rowcount As LongLong
Dim start As Double
Dim closer As Double
Dim price_change As Double
Dim percent_change As Double

  ' Set an initial variable for holding the total volume per ticker
  Dim Volume_Total As Double
  Volume_Total = 0

' ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

start = Cells(2, 3).Value

 'For Loop to iterate through whole sheet - Define rowcount
 rowcount = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To rowcount

'No change in ticker
If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
      
    ' Add and store Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
      
    'Store open value from day 1 as start

'When we reach a change in ticker
ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

    'Set ticker name
    Ticker = Cells(i, 1).Value


    'Do stuff for close data
    closer = Cells(i, 6).Value
    
    'Do calc for open-close
    price_change = closer - start
    percent_change = (closer - start) / start
    start = Cells(i + 1, 3).Value
    
 
    'Do volume calc
   
    'Output data to table
    
    ' Add and store Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
    
    ' Print the Ticker Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker
    
      ' Print the Price Change in the Summary Table
      Range("J" & Summary_Table_Row).Value = price_change
    
      ' Print the Percent Change in the Summary Table
      Range("K" & Summary_Table_Row).Value = percent_change
      
    ' Print the Total Volume in the Summary Table
      Range("L" & Summary_Table_Row).Value = Volume_Total
      
    ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Volume_Total = 0
      
End If

Next i



End Sub


