# Stock Analysis with VBA

## Overview of Project

Refactor existing code to analyze data for a number of stocks in a large database:
- Refactor previous macro to enhance proficiency
- Add Arrays to the macro
- Compare the time from the refactored macro to the previous macro

### Purpose

  The purpose of this project was to create and refactor database analysis macro in VBA for our client, Steve, to help him analyze a large stock database from multiple years of data as quickly as possible. This project included a **refactoring** of our original macro which included many ***For*** loops and ***If...Then*** statements. This required using an array approach to the project. We also included a **User Input** box and **Run Analysis** and **Clear Data** buttons to assist Steve in simplifying the user interface with his database file.
  
  Steve will be able to use our final product to assist his clients by providing historical stock performance information to help guide their future stock purchase decisions.

## Results

### Comparison of 2017 vs. 2018 Stock Performance

When we run an annual stock comparison between 2017 and 2018 for the 12 stocks in this database, a number of things stand out:
* Overall, 2017 seemed to be a very good year for many stocks as 11 of the 12 stocks posted positive returns in their annual price per share. One quarter of the stocks achieved gains between 100% to 200% in their price. The only stock posting a loss was TERP.
* On the other hand, 2018 wasn't as favorable to the stock prices of our 12 stock sample. 10 of 12 posted negative annual returns in their price per share. Not as significant as the positive returns of 2017, but some stocks lost around 60% of their value.
* There were 2 stocks that posted year over year gains, ENPH and RUN.
* Overall total volumes traded of each stock varied year to year with no apparent correlation between volumes and gains/losses.

### Comparsion of Execution Times of Original Macro and Refactored Macro

[Speed of 2017 analysis with refactored macro.](/Resources/VBA_Challenge_2017.png)

> The link above is a screen shot of the run time for one attempt of the 2017 stock data utilizing the refactored code. This represents about and 80% decrease in the run time between the original code and the refactored code when running the script on the 2017 data. The run time on the original code was about .617 seconds vs. the .121 seconds of the refactored code as indicated by the image linked above.

[Speed of 2018 analysis with refactored macro.](/Resources/VBA_Challenge_2018.png)

> The link above is a screen shot of the run time for one attempt of the 2017 stock data utilizing the refactored code. This represents about an 88% decrease in the run time between the original code and the refactored code when running the script on the 2018 data. The run time on the original code was about .629 seconds vs. the .078 seconds of the refactored code as indicated by the image linked above.

Much of this increase in code performance can be accounted for by a few changes in the original code:
1. Exporting the data to Arrays that were not in the original code. The lines of code that created the arrays are below.
```
  Dim tickerVolumes(12) As Long
  Dim tickerStartingPrices(12) As Single
  Dim tickerEndingPrices(12) As Single
```
2. Replacing some of the original loop coding that searched each line of the database and output the data at the end of each loop:
```
  For i = 0 To 11
          ticker = tickers(i)
          totalVolume = 0

          '5) loop through rows in the data
          Worksheets(yearValue).Activate
          For j = 2 To RowCount
              '5a) Get total volume for current ticker
              If Cells(j, 1).Value = ticker Then

                  totalVolume = totalVolume + Cells(j, 8).Value

              End If
              '5b) get starting price for current ticker
              If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

                  startingPrice = Cells(j, 6).Value

              End If
              '5c) get ending price for current ticker
              If Cells(j, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
              endingPrice = Cells(j, 6).Value

              End If

          Next j
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
```
with coding that would use array indexing to populate the arrays and output the data after the arrays have been populated:
```
''2b) Loop over all the rows in the spreadsheet.
    tickerIndex = 0
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
                
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
       
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
       
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If
        
    Next i
        
       
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```
## Summary

By completing this assignment, it became apparent that there are some advantages to refactoring code, which include:
1. A functional macro already exists to perform the task that you are looking to perform, thus saving some time in creating a more efficient one.
2. It is somewhat easy to see where the existing code can be adjusted by looking for the ***For*** loops and nested ***If...Then*** statements.
3. The refactored code usually required less commands in general by decreasing repetition within the original code. 
4. The new code is typically more efficient and performs it's function much quicker than the original code.

There are a couple disadvantages of refactoring code that I encoutered during this project:
1. It took some time to determine how to refactor some of the original code and adjust the commands to utilize arrays.
2. New bugs can be introduced which can require some time to debug and correct.

For this specific project, there was a definite advantage in macro execution time between the original code and the refactored code. This was discussed previously in the Results section above. Another advantage that I observed was that the refactored code was much cleaner and easier to follow as some of the loops and If...Then statments could be removed, thus making the final code cleaner. 

Some of the disadvantages that I encoutered was that it did take a fair amount of time to figure out what could be changed to improve the script performance. I attribute most of this to my unfamiliarity with VBA and coding as this is very new to me. During the refactoring process I also found that the basic functionality of the original code changed due to inconsistencies or errors in my refactored code, thus creating new bugs. This took some time to debug and get the new code to do what I needed it to do. So if time is a constraint, refactoring a properly functioning code may not be the best option.
