
Sub stocks__analysis()

'----------------------------------------
'---Set an initial variables

'define initial variables required
Dim ticker_name     As String '--Ticker
Dim total_stock     As Double '--Total Stock Volume
Dim open_year       As Double '--Open value for each year
Dim close_year      As Double '--Closevalue for each year
Dim change_year     As Double '--Change value  for each year
Dim change_perc     As Double '--Change in % for each year

'define parameters
Dim date_num        As Integer '--may need to use date as integer
Dim start_data      As Integer '--row start
Dim sum_table_row   As Integer
Dim ws              As Worksheet
'define worksheet
Dim worksheetname   As String

'----------------------------------------
'---Begin loop for all worksheets

For Each ws In Worksheets

worksheetname = ws.Name
'----------------------------------------
''------Create first Summary Table column 9 - 12

    'Create Headers for new table array
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'set up integer where loop starts
    'Set an initial variable Total_stoack for holding the total per Ticker Name

    sum_table_row = 2 ' avoid writing in header row
    previous_i = 1
    total_stock = 0

    'Count the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all stock volume for yearly change,percent change and total stock
    
        For i = 2 To lastrow

        ' Check if we are still within the Ticker name, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
        ' Set the Ticker name
        ticker_name = ws.Cells(i, 1).Value

        'move on to the next ticker
        previous_i = previous_i + 1
      
        open_year = ws.Cells(previous_i, 3).Value
        close_year = ws.Cells(i, 6).Value

        'a loop for total_stock in column 7
        For j = pervious_i To i

            ' Add to the Ticker Total
            total_stock = total_stock + ws.Cells(i, 7).Value
        
        Next j

        If open_year = 0 Then
            change_perc = close_year

        Else
            change_year = close_year - open_year
            change_perc = change_year / open_year
        End If

    '----------------------------------------

    'Print Output
      ws.Cells(sum_table_row, 9).Value = ticker_name
      ws.Cells(sum_table_row, 10).Value = change_year
      ws.Cells(sum_table_row, 11).Value = change_perc

      ws.Cells(sum_table_row, 11).NumberFormat = "0.00%" '--change column 11 to percent format

      ws.Cells(sum_table_row, 12).Value = total_stock

      sum_table_row = sum_table_row + 1
      ' Reset the Ticker Total
      total_stock = 0
      change_year = 0
      change_perc = 0
      ' If the cell immediately following a row is the same Ticker.
      previous_i = i
 
    End If

  Next i

'MsgBox ("First Summary Table Complete")

'----------------------------------------
''------Create Bonus  Summary Table column




'----------------------------------------
''------Insert conditional formatting - note % formate already completed in loop

    jlastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        For j = 2 To jlastrow
            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

'MsgBox ("SHEET " + worksheetname + " Complete")
Next ws

MsgBox ("Final Analysis Complete!!")

End Sub



