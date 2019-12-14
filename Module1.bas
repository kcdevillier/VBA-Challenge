Attribute VB_Name = "Module1"
Sub StockEvaluation()

'define variables
Dim ticker As String
Dim LastRow As Double
Dim open_ticker As String
Dim yearly_Change As Double
Dim percent_change As Double
Dim total_stock As Double
Dim stock_volume As Double
Dim total_cell_location As Integer
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double

'Start loop through worksheets
For Each ws In Worksheets
    
    'set LastRow length and reset variables
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1
    total_stock = 0
    total_cell_location = 2
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    'Set first open ticker value of worksheet
    open_ticker = ws.Cells(2, 3).Value

    'create new column headers for each worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'set cell style type for columns
    ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
    ws.Range("o2:o3").NumberFormat = "0.00%"
    
    'Create New column headers for greatest increase section
    ws.Range("n1").Value = "Ticker"
    ws.Range("o1").Value = "Value"
    
    'Adjust column width to fit text
    ws.Columns("J").ColumnWidth = 13
    ws.Columns("K").ColumnWidth = 18
    ws.Columns("L").ColumnWidth = 18
    ws.Columns("M").ColumnWidth = 19
    
    'Add Row names for greatest increase & decrease
    ws.Range("m2").Value = "Greatest % Increase"
    ws.Range("m3").Value = "Greatest % Decrease"
    ws.Range("m4").Value = "Greatest Total Volume"
     
    'Loop through rows
    For i = 2 To LastRow

        'store current ticker cell value into variable
        ticker = ws.Cells(i, 1).Value
        stock_volume = ws.Cells(i, 7).Value

        'If ticker values are the same then keep adding stock volume
        If (ticker = ws.Cells(i + 1, 1)) Then

            total_stock = total_stock + stock_volume

        'if ticker values are not equal then a different stock ticker is next
        'therefore ....
        ElseIf (ticker <> ws.Cells(i + 1, 1)) Then
            
            'Set Yearly Change and add into its appropiate cell location
            yearly_Change = ws.Cells(i, 6).Value - open_ticker
            ws.Cells(total_cell_location, 10).Value = yearly_Change
            
            'Last time that total_stock gets added before resetting to 0 for next ticker
            total_stock = total_stock + stock_volume
            
            'compare current stored stock volume to each tickers total volume
            'to find greatest volume from all sheets
            If (greatest_volume < total_stock) Then
            
                'set greatest volume and greatest_volume to appropiate cell
                greatest_volume = total_stock
                ws.Range("O4").Value = greatest_volume
                ws.Range("N4").Value = ticker
            
            End If
            
            'conditional formatting for yearly_change
            'if negative change set cell to red
            If (yearly_Change < 0) Then
                ws.Cells(total_cell_location, 10).Interior.ColorIndex = 3
                
            'if positive change set cell to green
            ElseIf (yearly_Change > 0) Then
                ws.Cells(total_cell_location, 10).Interior.ColorIndex = 4
                
            End If
            
            'if closing & opening value are at 0 then
            If (open_ticker = 0 And ws.Cells(i, 6).Value = 0) Then
            
                'fill appropiate cells with 0
                ws.Cells(total_cell_location, 11).Value = 0
                
            'if closing value is 0 but not opening then
            ElseIf (open_ticker = 0) Then
            
                'Find percentage
                percent_change = ((ws.Cells(i, 6).Value) / 100)
                ws.Cells(total_cell_location, 11).Value = percent_change
                
            Else
            
               'Set Percentage Change and add it to its appropiate cell location
                percent_change = (((((ws.Cells(i, 6).Value) / open_ticker) * 100) - 100) / 100)
                ws.Cells(total_cell_location, 11).Value = percent_change
            
            End If
            
            'calculate greatest increase and decrease
            If (greatest_decrease > percent_change) Then
            
                greatest_decrease = percent_change
                ws.Range("N3").Value = ticker
                ws.Range("O3").Value = greatest_decrease
                
            ElseIf (greatest_increase < percent_change) Then
            
                'set greatest_increase and place values in greatest increase cells
                greatest_increase = percent_change
                ws.Range("N2").Value = ticker
                ws.Range("O2").Value = greatest_increase
                
            End If
            
            'set new open_ticker value
            open_ticker = ws.Cells(i + 1, 3).Value
            
            'Total volume into its appropiate cell and add to cell location counter
            ws.Cells(total_cell_location, 12).Value = total_stock
            ws.Cells(total_cell_location, 9).Value = ticker
            total_cell_location = total_cell_location + 1
            
            'reset values for new ticker, percent_change & yearly_change
            total_stock = 0
            yearly_Change = 0
            percet_change = 0

        End If

    Next i
     
Next ws

End Sub

