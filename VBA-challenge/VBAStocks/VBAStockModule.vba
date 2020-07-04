Attribute VB_Name = "Module1"
Sub Code()

 For Each ws In Worksheets
 

''Column Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
   
    ws.Columns("I:L").AutoFit
    
   
    Dim TickerCol As Integer
    TickerCol = 9
   
    Dim TickerRow As Integer
    TickerRow = 2
   
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim TotalVolumeCol As Integer
    TotalVolumeCol = 12
    
    Dim OpeningValue As Double
    Dim ClosingValue As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    YearlyChangeCol = 10
    PercentChangeCol = 11
    
    
    Dim TotalVolume As Double
   
    'Pick out unique ticker symbols

    OpeningValue = ws.Cells(2, 3).Value
   
   For i = 2 To lastrow
   
        Dim CurrentCell As String
        CurrentCell = ws.Cells(i, 1).Value
       
        Dim NextCell As String
        NextCell = ws.Cells(i + 1, 1).Value
       
        If CurrentCell <> NextCell Then
    
        'Print unique ticker symbol starting in I2
        ws.Cells(TickerRow, TickerCol).Value = CurrentCell
        
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        'assign the value of the total volume to the value of the cells in the totalvol column"
          ws.Cells(TickerRow, TotalVolumeCol).Value = TotalVolume
       
            ClosingValue = ws.Cells(i, 6)
            
            
           YearlyChange = ClosingValue - OpeningValue
           
        If OpeningValue <> 0 Then
            

           PercentChange = YearlyChange / OpeningValue
           
           Else
           
           PercentChange = 0
           
           End If
           
          ws.Cells(TickerRow, YearlyChangeCol).Value = YearlyChange
           
           ws.Cells(TickerRow, PercentChangeCol).Value = PercentChange
                If YearlyChange < 0 Then
                ws.Cells(TickerRow, YearlyChangeCol).Interior.ColorIndex = 3
                Else
                    ws.Cells(TickerRow, YearlyChangeCol).Interior.ColorIndex = 4
                

                End If
                
           
           ws.Cells(TickerRow, PercentChangeCol).NumberFormat = "###,##0.00%"
           
           TickerRow = TickerRow + 1
           
           OpeningValue = Cells(i + 1, 3)
           
            TotalVolume = 0
        
        
        Else
    
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        PercentChange = 0
        
        
        End If
        

    Next i
    
    Next ws
     
    End Sub

