Sub StockAnalysis():

    'Disable screen updating until script is finished to improve performance
    Application.ScreenUpdating = False
    
    'Setup worksheet as variable for final looping
    Dim ws As Worksheet
    
    'Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        'Activate the current worksheet
        ws.Activate
        
        'On each worksheet do the following:
        'Set New Column Headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % decrease"
        Range("O4").Value = "Greatest total volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Define Variables
        Dim StockCount As Long
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Double
        Dim LastRow As Long
        Dim YearBegin As String
        Dim YearEnd As String
        Dim YearOpen As Double
        Dim YearClose As Double
        
        'Use the individual sheet names to store the year begin/end dates
        YearBegin = ws.Name & "0102"
        YearEnd = ws.Name & "1231"
        
        'Initialize variables that are used in first loop
        StockVolume = 0
        StockCount = 0
        
        'Find and store the Last Row in the Sheet
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through the tickers
        'This loop is to grab the unique ticker names and grab the necessary yearly data
        For i = 2 To LastRow
            
            
            'If its a new year than its a new ticker
            If Cells(i, 2) = YearBegin Then
                
                'grab the open value for the year
                YearOpen = Cells(i, 3).Value
                
                'grab the volume for this day
                StockVolume = StockVolume + Range("G" & i).Value
                
            'If the next 2 rows are the same ticker
            ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                
                'keep adding the volume
                StockVolume = StockVolume + Range("G" & i).Value
            
            'Inevitably you reach the end of the year
            Else:
                
                'Grab the final close value for the year
                YearClose = Cells(i, 6).Value
                
                'Add the unique ticker name to column I (StockCount+2 is to bypass the header row and to account for the fact that StockCount starts at 0)
                Range("I" & StockCount + 2).Value = Cells(i, 1).Value
                                
                'Calculate the yearly change
                YearlyChange = YearClose - YearOpen
                
                'Put the yearly change in the correct cell for the ticker
                Cells(StockCount + 2, 10).Value = YearlyChange
                
                'Math out percentage change
                'Rounding here will give percentage to 2 decimal places (testing for efficiency, maybe not necessary)
                PercentChange = Round(YearlyChange / YearOpen, 4)
                
                'Put the percentage change in the correct cell for the ticker
                Range("K" & StockCount + 2).Value = PercentChange
                
                'Format Percent change cells so that they display as a percent
                Range("K" & StockCount + 2).NumberFormat = "0.00%"
                
                'Calculate the final Stock volume
                StockVolume = StockVolume + Range("G" & i).Value
                
                'Put the final stock volume in the correct cell for the ticker
                Cells(StockCount + 2, 12).Value = StockVolume
                
                'Reinitialize stock volume fhr the next iteration
                StockVolume = 0
                
                'StockCount ends up being a count of the unique stock tickers at the end of the loop (technically StockCount+1 is the final number of unique tickers since StockCount starts at 0)
                StockCount = StockCount + 1
                
            End If
                        
        Next i
        
        
        'Conditional Formatting--------------------------------------------------------------
        'Note intentionally leaving a change of "0" unformatted since 0 is neither positive or negative
            
        For i = 2 To StockCount + 1
            
            If Range("J" & i).Value < 0 Then
                
                Range("J" & i).Interior.ColorIndex = 3
                Range("K" & i).Interior.ColorIndex = 3
                
            ElseIf Range("J" & i).Value > 0 Then
            
                Range("J" & i).Interior.ColorIndex = 4
                Range("K" & i).Interior.ColorIndex = 4
                
            End If
            
        Next i
        
    'Getting the min/maxes from the data------------------------------------------------------
        
        'Declare variables
        Dim MinChange As Double
        Dim MaxChange As Double
        Dim MaxVolume As Double
        
        'Store the min/max of the quantities
        MaxDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & StockCount + 2))
        MaxIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & StockCount + 2))
        MaxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & StockCount + 2))
        
        'Loop to identify the stock ticker associated with the min/max values
        For i = 2 To StockCount + 1
        
            If Range("K" & i).Value = MaxDecrease Then
                
                MaxDecreaseTicker = Range("I" & i).Value
                
            ElseIf Range("K" & i).Value = MaxIncrease Then
            
                MaxIncreaseTicker = Range("I" & i).Value
                
            ElseIf Range("L" & i).Value = MaxVolume Then
            
                MaxVolumeTicker = Range("I" & i).Value
                
            End If
            
        Next i
                
        'Put all of the found values in the sheet and format where necessary
        Range("P2").Value = MaxIncreaseTicker
        Range("P3").Value = MaxDecreaseTicker
        Range("P4").Value = MaxVolumeTicker
        Range("Q2").Value = MaxIncrease
        Range("Q3").Value = MaxDecrease
        Range("Q4").Value = MaxVolume
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        
    'Final Formatting--------------------------------------------------------------------------
    'Autofit column width (mainly to prevent scientific notation of volume)
    ws.Columns.AutoFit
        
    
    'jump to the next worksheet and do all of the above
    Next ws

    'Reactivate screen updating (for efficiency)
    Application.ScreenUpdating = True
    
    'jump to top of sheet when done because it was ending scrolled to bottom
    ActiveWindow.ScrollRow = 1

End Sub




