# Green Stock Analysis With VBA

## Overview of Project 
In this scenario, we have been assisting a client and his family with an analysis of green stock market options. We do this in the hopes of helping them make a decision with future investments. The available data shows various stocks, their daily value over a year's parameter, and the volumes associated with each stock - which in turn indicates the condition of a particular option.  
Our client would now like to analyze the data across any year that is available. The client still wants to know how well the stocks did from 2017 to 2018. However, we will change our script so that the client can carry out their own analysis with any datasheet regarding these stocks.

### Purpose
As such, we want to show how well these stocks performed between 2017 and 2018. We will alter the VBA script that was initially created to present any year's stock analysis for the client. The objective is to create a more efficent universal script that can be used for any year, and show the client just how good each stock is doing. 

## Results 
After running the refactored script and designating the year that was wanted, we saw an output presenting the Return for each stock of interest in 2017 and 2018, respectively. In 2017, all stock options except TERP showed a positive return percentage. The stock options DQ, ENPH, FSLR, and SEDG showed significant promise, with return percentages over one-hundred. 
Historically, the year 2018 marked considerable difficulties and downtrends in the stock market. The data for 2018 reflects that trend. All stock options of interest except ENPH and RUN had negative return percentages. 

In order to show our client the efficiency of the new script, we have shown portions of the previous script and refactored script - along with their associated execution times. Overall, it is clear that refactoring the script allowed for a much shorter execution time. 
Using the refactored script, execution times for either year were below 0.36 seconds. On the other hand, the original script execution time for 2018 was approximately 1.5 seconds. This comparison is a great example depicting the efficiency of the refactored script. 

### Refactored Script 
As requested by the client, we have compared the execution times of the refactored script against the original script; and displayed examples of script along with the associated images. 
 
	
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
            
                '3d) Increase the tickerIndex.
               
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

![Refactored_Analysis_for_2017_&_Run_Time](https://github.com/JV348/stock-analysis/blob/211858ae7111e37e8e6470935c61b562ffc21965/Resources/VBA_Challenge_2017.png)
![Refactored_Analysis_for_2018_&_Run_Time](https://github.com/JV348/stock-analysis/blob/211858ae7111e37e8e6470935c61b562ffc21965/Resources/VBA_Challenge_2018.png)


### Patterns From Previous Script

  'Set initial volume to zero
    totalVolume = 0

  'Establish the number of rows to loop over
    rowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'Loop over all the rows
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

 '4) Loop through tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
          '5) loop through rows in the data
           Worksheets("2018").Activate
           For j = 2 To RowCount
          '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                
              totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
           '5b) Get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                startingPrice = Cells(j, 6).Value
                
            End If
            
            '5c) Get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
    
            
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
        endTime = Timer
        MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)

![Original_Script_&_Run_Time](https://github.com/JV348/stock-analysis/blob/ce0ef032166ac698d6ec3bbb585f7a094f3be7d1/Resources/Original_script.png)
 	  
## Summary

- What are the advantages or disadvantages of refactoring code?

As noted before, a great advantage to refactoring the code was the fact that the execution time was much shorter. Furthermore, this instance of refactoring simplified the manner by which the intended user could interact with multiple data. It was also worth noting that the length of the actual script was shorter and easier for another developer to integrate into their projects. Above all, the refactored script was effective at looking through any and all of the data that was available.
Unfortunately, it is possible that without an adequate explanation of the Macros being executed, it can be easy for another developer to get lost. The original scripts within the module were very particular and specific to a worksheet. The refactored script is rather generic and lacks some of that explanation. Also, the refactored script was actually more difficult to debug. Small errors always affect the output of a script, but these difficulties were more pronounced with the new script. 

- How do these pros and cons apply to refactoring the original VBA script?

Refactoring the VBA script led to greater efficiency, as shown by the differences in execution time. The refactored script was also shortened in length significantly, considering it utilized only some patterns found in the original script. 
On the other hand, some of the new script may confuse a fellow developer who is not familar with the entire process that has been executed. And of course, looping through all data can lead to some frustrating debugging issues if a developer makes an error in any given line of script. 