Sub stocksummary()
    
    'Loop through all worksheets in the workbook - added to move across all sheets
    
        Dim ws                  'not sure why this is needed
        Application.StatusBar = "Please Wait"
        Application.Cursor = xlWait
        Application.ScreenUpdating = False
        
    For Each ws In Worksheets
         
        'label and format the Summary Table headers and format colum K for % and
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("K:K").NumberFormat = "0.0%;[Red]-0.0%"
        ws.Range("L:L").NumberFormat = "#,###"
        ws.Range("I1:L1").Font.Bold = True
        ws.Range("I1:L1").Font.Underline = True
        ws.UsedRange.EntireColumn.AutoFit
    
       'sort tables in ticker code order as check
        Range("A1").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Sort key1:=Range("A1", Range("A1").End(xlDown)), _
        order1:=xlAscending, Header:=xlYes
        Range("A1").Select
        
        'establish variables (moved to Long to fix bug)
        Dim tickercode As String
        Dim vol As Double
        Dim sum_tickercodes As Long
        Dim closeVal As Double
        Dim yearlyChge As Double
        Dim percentChge As Double
        Dim LastRow As Long
        Dim i As Long
        Dim openVal As Double
        Dim lastrow_new_table As Long
        
        'set variable start, count etc.
        vol = 0
        sum_tickercodes = 2                                                         '>> tickercode start point
        openVal = Cells(2, 3).Value                                               '>>required to set initial opening valuation
        Const StartRow As Byte = 2

        'set the data range being managed ************************
        LastRow = ws.Range("A" & StartRow).End(xlDown).Row
        
        'loop through the rows by the ticker codes
        
        For i = 2 To LastRow

            'If statement to determine when ticker codes change
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'capture ticker codes
              tickercode = ws.Cells(i, 1).Value

              'ticker codes to summary table
              ws.Range("I" & sum_tickercodes).Value = tickercode
              
              'add the volume of trade
              vol = vol + ws.Cells(i, 7).Value
             
              'trade vol in summary table
              ws.Range("L" & sum_tickercodes).Value = vol

              'closing price capture
              closeVal = ws.Cells(i, 6).Value

              'yearly Change calculation (close-open)
              yearlyChge = (closeVal - openVal)
              
              'yearly change in summary table
              ws.Range("J" & sum_tickercodes).Value = yearlyChge

             'percent change calculation
              percentChge = yearlyChge / openVal
                  
              'yearly change in the summary table
              ws.Range("K" & sum_tickercodes).Value = percentChge
              
              'reset the row counter. Add one to the summary_ticker_row
              sum_tickercodes = sum_tickercodes + 1

              'reset trade vol to clear
              vol = 0

              'reset the openVal to clear
              openVal = ws.Cells(i + 1, 3)
            
            Else
              
            'Add the volume of trade
             vol = vol + ws.Cells(i, 7).Value

            End If
        
        Next i

    'Conditional formatting that will highlight positive change in green and negative change in red

    lastrow_new_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To lastrow_new_table
            If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
    
    Next i
'--------------------------------Bonus----------------------------------

'create, label and format the summary Greatest Increase and Decrease sections
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume "
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Range("o1:p1").Font.Bold = True
        ws.Range("N2:N4").Font.Bold = True
        
    For i = 2 To lastrow_new_table
    
    'find the greatest percent change and format number
    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_new_table)) Then
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value   'ticker code
                    ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).NumberFormat = "0.00%"

            'find the greatest decrease percent change and format number
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_new_table)) Then
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value  'ticker code
                    ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).NumberFormat = "0.00%;[Red]-0.00%"
                    
            
            'find the greatest total volume and format number
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_new_table)) Then
                    ws.Cells(4, 15).Value = (ws.Cells(i, 9).Value)  'ticker code
                    ws.Cells(4, 16).Value = (ws.Cells(i, 12).Value)
                    ws.Cells(4, 16).NumberFormat = "#,###"
    
            'autofit the contents to capture Vol
            ws.UsedRange.EntireColumn.AutoFit
    
    End If
    
    Next i
    
    Next ws
    
        Application.StatusBar = ""
        Application.Cursor = xlDefault
        Application.ScreenUpdating = True
    
End Sub




