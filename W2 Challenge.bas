Attribute VB_Name = "Module1"
'Step 1 Create Ticker List & Total Stock Column

Sub Ticker_list()

For Each ws In Worksheets

Dim worksheetname As String
  
  'Count the number of rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set up Worksheet variable
worksheetname = ws.Name

  ' Set an initial variables
    Dim Ticker_name As String
    Dim date_num As Integer
  ' Set an initial variable Total_stoack for holding the total per Ticker Name
  Dim Total_stock As Double
  Total_stock = 0
  
  ' Keep track of the location for each Ticker Name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Stock Volume
  For i = 2 To LastRow
    ' Check if we are still within the Ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ' Set the Brand name
      Ticker_name = ws.Cells(i, 1).Value
      ' Add to the Brand Total
      Total_stock = Total_stock + ws.Cells(i, 7).Value
      
    'Print Output
      ws.Range("I" & Summary_Table_Row).Value = Ticker_name
      ws.Range("L" & Summary_Table_Row).Value = Total_stock

      Summary_Table_Row = Summary_Table_Row + 1
      ' Reset the Brand Total
      Total_stock = 0
    ' If the cell immediately following a row is the same brand...
    Else
      ' Add to the Brand Total
      Total_stock = Total_stock + ws.Cells(i, 7).Value
    End If
  Next i

' Create Headers for new table array

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"


Next ws

MsgBox ("Run Complete")

End Sub

'Step 2 Convert date to number to find min&max
'Use Date and Ticker to create Change and % Change



Sub ConvertTextToNumber()

[B:B].Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

End Sub


'Step 3 Set up formatting

Sub Colour_format()
For Each ws In Worksheets
Dim worksheetname As String
lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
worksheetname = ws.Name

Dim k As Double
Dim Percent_Change As Double

For k = 2 To lastrow2
    If (Percent_Change >= 0) Then
        ws.Cells(k, 11).Interior.ColorIndex = 4
    
    ElseIf (Percent_Change < 0) Then
        ws.Cells(k, 11).Interior.ColorIndex = 3
    End If
Next k
Next ws
End Sub
