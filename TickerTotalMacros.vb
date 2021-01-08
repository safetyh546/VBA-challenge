Sub ClearContents():
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Columns(9).ClearContents
    ws.Columns(9).ClearFormats
    ws.Columns(10).ClearContents
    ws.Columns(10).ClearFormats
    ws.Columns(11).ClearContents
    ws.Columns(11).ClearFormats
    ws.Columns(12).ClearContents
    ws.Columns(12).ClearFormats
Next ws
End Sub



Sub TickerTotal():

Dim ws As Worksheet

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
       
        Dim Total As Double
        Dim Ticker As String
        Dim TickerCounter As Integer
        Dim TickerOpenPrice As Double
        Dim TickerClosePrice As Double
        
        'set Yearly change column format to percent with 2 decimal places
        ws.Columns(11).NumberFormat = "0.00%"
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set starting values for Total and TickerCounter variables
        Total = 0
        TickerCounter = 1

        'set first open price of ticker
        TickerOpenPrice = ws.Cells(2, 3).Value
        
        'set column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'loop througha all data rows and set totals in columns 9-12
        For r = 2 To LastRow
            ' set initial values from first row of ticker
            Ticker = ws.Cells(r, 1).Value
            Total = Total + ws.Cells(r, 7).Value
                 
            
            ' do action when you are on last row of Ticker
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                     
               'Increment counter and grab closing price
               TickerCounter = TickerCounter + 1
               TickerClosePrice = ws.Cells(r, 6).Value
               
               'Print Ticker
               ws.Cells(TickerCounter, 9).Value = Ticker
               
               'Print Yearly Change
               ws.Cells(TickerCounter, 10) = TickerClosePrice - TickerOpenPrice
               
               If TickerClosePrice - TickerOpenPrice > 0 Then
                      ' Set the Cell Colors to Green
                      ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                      ' Set the Cell Colors to Red
               Else: ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
               End If
               
               'Print Percent Change
               'if open price is zero, cannot calc so put N/A
               If TickerOpenPrice <> 0 Then
                  ws.Cells(TickerCounter, 11) = (TickerClosePrice - TickerOpenPrice) / TickerOpenPrice
               Else: ws.Cells(TickerCounter, 11) = "N/A"
               End If
                             
               'Print Total Stock Volume
               ws.Cells(TickerCounter, 12) = Total
        
               ' reset total for next ticker
               Total = 0
               
               'Set next Ticker open price
               TickerOpenPrice = ws.Cells(r + 1, 3).Value
        
             End If
            Next r
        
        
        'Autofit Ticker total columns
        ws.Columns("I:L").AutoFit
        
    ' --------------------------------------------
    ' FIXES COMPLETE
    ' --------------------------------------------
    Next ws
    
End Sub








