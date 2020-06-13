Attribute VB_Name = "Module1"
Sub StockSummary()

Dim ws As Worksheet 'worksheet variable
Dim t, NextRow As Integer 't is for Ticker column & NextRow is to print new ticker on summary tabel
Dim YearCh, PerCh, TotalStock As Double
Dim Y_Open, Y_Close As Double
Dim Row As Double
Dim Ticker As String

'loop through entire worksbook for each worksheet in workbook
For Each ws In Worksheets

    'To add header for all of the summay table
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    NextRow = 2 'Initial start row cell value for the summary table
        
   'initial values for the ticker and year open for the perticual ticker
    Ticker = ws.Cells(2, 1).Value
    Y_Open = ws.Cells(2, 3).Value
   
   '-------------------------------------------------------------------------------------------------------------
   
    'Calculate the number of populated rows plus one to ensure last ticker is populated in summary table
    Row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    'loop through worksheet to find all diffrent tickets
    For t = 2 To Row
        If ws.Cells(t, 1).Value = Ticker Then
                        
            TotalStock = TotalStock + ws.Cells(t, 7).Value
                                     
        Else
           
                Y_Close = ws.Cells(t - 1, 6).Value 'Year close value from last ticker for each ticker
            
                'addressing Year close value being 0
                If Y_Open = 0 Then
                    YearCh = Y_Close - Y_Open
                    PerCh = 0
                Else
                    'Yearly Change and Percent change calculation
                    YearCh = Y_Close - Y_Open
                    PerCh = (YearCh / Y_Open)
                End If
            
               'adding values to the summary table
               ws.Cells(NextRow, 9).Value = Ticker
               
               ws.Cells(NextRow, 10).Value = YearCh
                ws.Cells(NextRow, 10).NumberFormat = "$#,##0.00" 'Two Decimal formating
                    
                    'formating to highlight cells in green or red
                    If ws.Cells(NextRow, 10).Value <= 0 Then
                        ws.Cells(NextRow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(NextRow, 10).Interior.ColorIndex = 4
                    End If
               
               ws.Cells(NextRow, 11).Value = PerCh
                ws.Cells(NextRow, 11).NumberFormat = "0.00%" 'percent formating
                
               ws.Cells(NextRow, 12).Value = TotalStock
                                            
            NextRow = NextRow + 1 'moving to next row on summary table
            Ticker = ws.Cells(t, 1).Value 'intialize the ticker value to next ticker in the table
            Y_Open = ws.Cells(t, 3).Value 'intizlize the Y_Open value to next ticker year open value
            TotalStock = ws.Cells(t, 7).Value 'TotalStock volume is assigned back to the cell where t is currently in the loop to start adding until ticker values are not same
            
        End If
                                      
    Next t
    
    '----------------------------------------------------------------------------------------------------------------
 
    Dim YCRow As Double
    
    'To add header for all of the summay table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    YCRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
   
    '--------------------------------------------------------------------------------------------------------------
    
    Dim Max, Min, HiVol, VolRow As Double
    Dim tMax, tMin, tV As Integer
    Dim TiK, TiKV, TiKM As String
    
    Max = 0
    TiK = ""
    'Loop to identify Max percent increase
    For tMax = 2 To YCRow
        If ws.Cells(tMax, 11).Value > Max Then
            Max = ws.Cells(tMax, 11).Value
            TiK = ws.Cells(tMax, 9).Value
        End If
    Next tMax
    
    ws.Cells(2, 16).Value = TiK
    ws.Cells(2, 17).Value = Max
        ws.Cells(2, 17).NumberFormat = "0.00%"
    
   '------------------------------------------------------------------------------------------------------------
    Min = 0
    TiKM = ""
    
    'Loop to identify Min percent increase
    For tMin = 2 To YCRow
         If ws.Cells(tMin, 11).Value < Min Then
            Min = ws.Cells(tMin, 11).Value
            TiKM = ws.Cells(tMin, 9).Value
        End If
    Next tMin
    
    ws.Cells(3, 16).Value = TiKM
    ws.Cells(3, 17).Value = Min
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
    '-----------------------------------------------------------------------------------------------------------
    HiVol = 0
    TiKV = ""
    
    VolRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    'Loop to identify Max Total Volume
    For tV = 2 To VolRow
        If ws.Cells(tV, 12).Value > HiVol Then
            HiVol = ws.Cells(tV, 12).Value
            TiKV = ws.Cells(tV, 9).Value
        End If
    Next tV
    
    ws.Cells(4, 16).Value = TiKV
    ws.Cells(4, 17).Value = HiVol
    
Next

End Sub
