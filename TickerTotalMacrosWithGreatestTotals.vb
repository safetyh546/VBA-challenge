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
    ws.Columns(17).ClearContents
    ws.Columns(17).ClearFormats
    ws.Columns(18).ClearContents
    ws.Columns(18).ClearFormats
    ws.Columns(19).ClearContents
    ws.Columns(19).ClearFormats
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
        Dim YearlyChange As Double
        Dim PctChange As Double
        Dim GreatestPctInc As Double
        Dim GreatestPctDecrease As Double
        Dim GreatestVolume As Double
        
        
        'set Percent change column format to percent with 2 decimal places
        ws.Columns(11).NumberFormat = "0.00%"
        ws.Range("s2").NumberFormat = "0.00%"
        ws.Range("s3").NumberFormat = "0.00%"
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set starting values for Total and TickerCounter variables
        Total = 0
        TickerCounter = 1
        GreatestPctInc = 0
        GreatestPctDecrease = 0
        GreatestVolume = 0

        'set first open price of ticker
        TickerOpenPrice = ws.Cells(2, 3).Value
        
        'set column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = "Greatest Total Volume"
        ws.Cells(1, 18).Value = "Ticker"
        ws.Cells(1, 19).Value = "Value"
        
        'loop througha all data rows and set totals in columns 9-12
        For r = 2 To LastRow
            ' set initial values from first row of ticker
            Ticker = ws.Cells(r, 1).Value
            Total = Total + ws.Cells(r, 7).Value
                 
            
            ' do action when you are on last row of Ticker
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                     
               'Increment counter and set values to variables
               TickerCounter = TickerCounter + 1
               TickerClosePrice = ws.Cells(r, 6).Value
               YearlyChange = TickerClosePrice - TickerOpenPrice
               If TickerOpenPrice <> 0 Then
                    PctChange = YearlyChange / TickerOpenPrice
               End If
               
               'Print Ticker
               ws.Cells(TickerCounter, 9).Value = Ticker
               
               'Print Yearly Change
               ws.Cells(TickerCounter, 10) = TickerClosePrice - TickerOpenPrice
               
               If YearlyChange > 0 Then
                      ' Set the Cell Colors to Green
                      ws.Cells(TickerCounter, 10).Interior.ColorIndex = 4
                      ' Set the Cell Colors to Red
               Else: ws.Cells(TickerCounter, 10).Interior.ColorIndex = 3
               End If
               
               'Print Percent Change and Greatest % Inc/Dec
               'if open price is zero, cannot calc so put N/A
               If TickerOpenPrice <> 0 Then
                  
                  'Percent change
                  ws.Cells(TickerCounter, 11) = PctChange
                  
                  'Greatest % Increase
                  If PctChange > GreatestPctInc Then
                       GreatestPctInc = PctChange
                       ws.Range("R2").Value = Ticker
                       ws.Range("S2").Value = PctChange
                  End If
                  
                  'Greatest % Decrease
                  If PctChange < GreatestPctDecrease Then
                       GreatestPctDecrease = PctChange
                       ws.Range("R3").Value = Ticker
                       ws.Range("S3").Value = PctChange
                  End If
               Else: ws.Cells(TickerCounter, 11) = "N/A"
               End If
                             
               'Print Total Stock Volume
               ws.Cells(TickerCounter, 12) = Total
               
               'Print Greatest Total Volume
               If Total > GreatestVolume Then
                    GreatestVolume = Total
                    ws.Range("R4").Value = Ticker
                    ws.Range("S4").Value = Total
               End If
               
         
               ' reset total for next ticker
               Total = 0
               
               'Set next Ticker open price
               TickerOpenPrice = ws.Cells(r + 1, 3).Value
        
             End If
            Next r
        
        
        'Autofit Ticker total columns
        ws.Columns("I:S").AutoFit
        
    ' --------------------------------------------
    ' FIXES COMPLETE
    ' --------------------------------------------
    Next ws
    
    MsgBox ("Ticker Totals Complete")
    
End Sub
