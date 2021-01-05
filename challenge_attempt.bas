Attribute VB_Name = "Module1"
Sub stock_data():
    
    Dim ws As Worksheet
    For Each ws In Worksheets

        ' Set variables
        Dim Ticker As String
        Dim rowcount As LongLong
        Dim start As Double
        Dim closer As Double
        Dim price_change As Double
        Dim percent_change As Double
        Dim yearly_change As Range
        Dim Volume_Total As Double
  
        'Set initial volume
        Volume_Total = 0

        ' Keep track of the location for each ticker in the summary table
         Dim Summary_Table_Row As Integer
         Summary_Table_Row = 2

        start = Cells(2, 3).Value
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"

         'For Loop to iterate through whole sheet - Define rowcount
         rowcount = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To rowcount

        'No change in ticker
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
      
      ' Add and store Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
       

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
        
        ' Add and store Volume Total
          Volume_Total = Volume_Total + Cells(i, 7).Value
        
        ' Print the Ticker Name in the Summary Table
          Range("I" & Summary_Table_Row).Value = Ticker
    
        ' Print the Price Change in the Summary Table
        Range("J" & Summary_Table_Row).Value = price_change
        Set yearly_change = Range("J" & Summary_Table_Row)
        yearly_change.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
          Formula1:="=0"
        yearly_change.FormatConditions(1).Interior.Color = vbRed
        yearly_change.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
          Formula1:="=0"
        yearly_change.FormatConditions(2).Interior.Color = vbGreen
    
          ' Print the Percent Change in the Summary Table
          Range("K" & Summary_Table_Row).Value = percent_change
          Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
          
        ' Print the Total Volume in the Summary Table
          Range("L" & Summary_Table_Row).Value = Volume_Total
          
        ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
      
         ' Reset the Brand Total
         Volume_Total = 0
      
         End If

         Next i

         Next ws
         
End Sub
