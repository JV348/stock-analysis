# Stock-Analysis

## Overview of Project 

### Purpose

## Results 

### Refactored Script 
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
 '1a) Create a ticker Index
        tickerIndex = 0
      
        '1b) Create three output arrays
        Dim tickerVolumes(11) As Long
        Dim tickerStartingPrices(11) As Single
        Dim tickerEndingPrices(11) As Single
    
    
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
       
        tickerVolumes(i) = 0
        
        Next i
            ''2b) Loop over all the rows in the spreadsheet.
             For i = 2 To RowCount
    
                '3a) Increase volume for current ticker
                If Cells(i, 1).Value = tickers(tickerIndex) Then
                
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
                End If
            
                '3b) Check if the current row is the first row with the selected tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
                End If
        
                '3c) check if the current row is the last row with the selected ticker
                'If the next row's ticker doesn't match, increase the tickerIndex.
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

                '3d Increase the tickerIndex.
               
                    tickerIndex = tickerIndex + 1
            
                End If
    
             Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
        Next i
![Refactored_Analysis_for_2017](https://github.com/JV348/stock-analysis/blob/211858ae7111e37e8e6470935c61b562ffc21965/Resources/VBA_Challenge_2017.png)
![Refactored_Analysis_for_2018](https://github.com/JV348/stock-analysis/blob/211858ae7111e37e8e6470935c61b562ffc21965/Resources/VBA_Challenge_2018.png)



### Previous Script

  'set initial volume to zero
    totalVolume = 0
  'Establish the number of rows to loop over
    rowStart = 2
    'DELETE: rowEnd = 3013
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd
    
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value
    
        End If
        
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        
            startingPrice = Cells(i, 6).Value
        
        End If
        
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        
            startingPrice = Cells(i, 6).Value
        
        End If
        
    Next i
 	  
## Summary

### Advantages or disadvantages of refactoring code

### Pros and cons to refactoring the original VBA script