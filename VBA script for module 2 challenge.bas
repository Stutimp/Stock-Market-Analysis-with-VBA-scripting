Attribute VB_Name = "Module1"
Sub Stockanalysis()
 Dim ws As Worksheet
 Dim ticker As String
 Dim nextticker As String
 Dim summaryrow As Double
 Dim openrate As Double
 Dim closerate As Double
 Dim YearlyChange As Double
 Dim LastRow As Long
 Dim PercentageChange As Double
 Dim totalVolume As Double
 Dim greatestIncrease As Double
 Dim greatestDecrease As Double
 Dim greatestTotalVolum As Double
 Dim greatestIncreaseTicker As String
 Dim greatestDecreaseTicker As String
 Dim greatesttotalVolumeTicker As String



'Initialize the variables for the Greatest values
greatestIncrease = 0
greatestDecrease = 0
greatestTotalVolume = 0


'To look Through sheets
    For Each ws In Worksheets
        summaryrow = 2
  

    'Determine the Last row in column A
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Initialize open rate and totalVolume for the first unique ticker
        openrate = ws.Cells(2, 3).Value
        totalVolume = 0
                
        'looping through data on the sheet
        For Row = 2 To LastRow
            ticker = ws.Cells(Row, 1).Value
            nextticker = ws.Cells(Row + 1, 1).Value
           
                ' find the last row for each ticker
            If nextticker <> ticker Then
                  
                    
            'Output the ticker symbol in Column J
                            
                  ws.Cells(summaryrow, 10).Value = ticker
                    
                'Adding name for the column
                  ws.Range("J1").Value = "Ticker"
                
            'Calculate yearly change for the current ticker
            
                  closerate = ws.Cells(Row, 6).Value
            
                  YearlyChange = closerate - openrate
            
        'Output the yearly change in column k for the current ticket and also applying conditional formating inside yearly change column k, with green for positive change and red for negative change
                        
                  ws.Cells(summaryrow, 11).Value = YearlyChange
                  If (YearlyChange > 0) Then
                     ws.Cells(summaryrow, 11).Interior.ColorIndex = 4
                  Else
                     ws.Cells(summaryrow, 11).Interior.ColorIndex = 3
                  End If
                
                                    
        'Adding the column header for the yearly change
                 ws.Cells(1, 11).Value = "Yearly change"
                         
        'calculate the percentage change for the current ticker
                 If openrate <> 0 Then
                 PercentageChange = (YearlyChange / openrate)
                 Else
                 PercentageChange = 0
                 End If
            
            'Now output the percentage change in column l for current ticker
            
                 ws.Cells(summaryrow, 12).Value = PercentageChange
                 ws.Cells(summaryrow, 12).Value = Format(PercentageChange, "0.00%")
            
            'Adding the column header for the percentage change
            
                 ws.Cells(1, 12).Value = "Percent Change"
                 
                 
            'Also applying conditional formating inside Percentage change column L, with green for positive change and red for negative change
            
                 If (PercentageChange > 0) Then
                 ws.Cells(summaryrow, 12).Interior.ColorIndex = 4
                 Else
                 ws.Cells(summaryrow, 12).Interior.ColorIndex = 3
                 End If
            'Output the total stock volume for in column M for the current ticker
                 totalVolume = totalVolume + ws.Cells(Row, 7).Value
                 
                 ws.Cells(summaryrow, 13).Value = totalVolume
                
            ' Adding the column header for the Total Stock Volume
                 ws.Cells(1, 13).Value = "Total stock Volume"
                      
                
                'Check for the "greatest values"
                 If PercentageChange > greatestIncrease Then
                    greatestIncrease = PercentageChange
                    greatestIncreaseTicker = ticker
                 ElseIf PercentageChange < greatestDecrease Then
                    greatestDecrease = PercentageChange
                    greatestDecreaseTicker = ticker
                 End If
                
                 If totalVolume > greatestTotalVolume Then
                        greatestTotalVolume = totalVolume
                        greatesttotalVolumeTicker = ticker
                 End If
               'reseting the total volume for the new ticker
                 totalVolume = 0
            
            'Moving to next row and update openrate and total Volume
                        
                 summaryrow = summaryrow + 1
                 openrate = ws.Cells(Row + 1, 3).Value
             
             Else
                'Accumulate the total volume for the current ticker
                 totalVolume = totalVolume + ws.Cells(Row, 7).Value
               
                      
    
            End If
        
        Next Row
       
        
        ' out put the "greatest" values and their respective tickers in all sheets

        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(2, 17).Value = greatestIncreaseTicker
        ws.Cells(2, 18).Value = greatestIncrease
        
        'to display percentage sign
        ws.Cells(2, 18).Value = Format(greatestIncrease, "0.00%")
        ws.Cells(1, 17).Value = "Ticker"
        
                
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 17).Value = greatestDecreaseTicker
        ws.Cells(3, 18).Value = greatestDecrease
        
        'To display percentage sign
        ws.Cells(3, 18).Value = Format(greatestDecrease, "0.00%")
        ws.Cells(1, 18).Value = "Value"
                
        
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(4, 17).Value = greatesttotalVolumeTicker
        ws.Cells(4, 18).Value = greatestTotalVolume
        
        greatestIncrease = 0
        greatestDecrease = 0
        greatestTotalVolume = 0
    Next ws
         
              
     

End Sub



