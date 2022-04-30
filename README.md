# Green Stock Analysis Using VBA

## Project Overview
For this project, we were given trading data for 12 different green energy companies (ie: opening and closing prices, and the volume of stocks sold on a daily basis). The data is divided into two sets of daily trade info for the years 2017 and 2018.

### Purpose
The purpose of this project was to refactor a previously written VBA macro that was created to run analysis on the trading volume and the return on investment (ROI) for the aforementioned stocks. The result of the analysis should inform which green energy companies are the best investments, based on how the stock value grew or fell over time.

The initial VBA code was written as part of the Module 2 VBA segments. The deliverable for the Module 2 challenge is to sucessfully refactor the code in a way that makes it run more efficiently, using the vbs code template provided as a starting place. Streamlining code to operate more efficiently is critical when you are working with large data sets.

## Results
Green energy stock had a positive ROI overall in 2017 and a negative ROI in 2018.

### 2017 Results 
In 2017 all of the stocks, minus TERP show a positive return on investment. DQ, ENPH, FSLR, and SEDG show the most significant positive returns. 

![2017 Performance](https://user-images.githubusercontent.com/103781847/166119164-b80191c1-d266-4d18-8d58-cc83d45847b4.png)

### 2018 Results 
In 2018, however, only two stocks showed a positive return: ENPH and RUN. The rest of the stocks had a negative return on investment. If I were advising our client, I would suggest that they start by investing in ENPH as the valuation of the other companies seems more volatile and risky. I would also advise them to wait to invest in other green energy companies until we have more confidence in the market for that industry. 

![2018 Performance](https://user-images.githubusercontent.com/103781847/166119168-f4b1e2ab-f0c7-4296-b424-67739c7e8dfb.png)

### Execution time for original code and refactored code 
By refactoring our code we did see a significant reduction in run time. For the 2017 data, we went from a run time of 0.265625 seconds to a run time of 0.0703125 seconds. For 2018, we reduced our original run time of 0.2734375 to 0.07421875 seconds. 

### 2017 Run Time Before and After

<img width="286" alt="VBA_Challenge_2017(pre-refactor)" src="https://user-images.githubusercontent.com/103781847/166118924-c801ff98-4fa7-4760-9260-7bca206da615.png">

![VBA_Challenge_2017](https://user-images.githubusercontent.com/103781847/166118935-58fddbbf-7208-4c8c-93e7-bd11adddce94.png)

### 2018 Run Time Before and After

<img width="283" alt="VBA_Challenge_2018(pre-refactor)" src="https://user-images.githubusercontent.com/103781847/166118960-6d74a5a4-37a5-40bd-a7ee-1b915f720ad9.png">

![VBA_Challenge_2018](https://user-images.githubusercontent.com/103781847/166118967-b81ad128-5eb5-41e6-b676-6dc460762fb8.png)

My hypothesis is that the delta on pre and post refactoring run-time differences would increase if our data set increases, so the refactoring will pay off if we need to work with much larger data sets. 

### Refactored macro
Below, I have included the main section of refactored code. To review the full orginal macro, and the refactored macro, please reference the .xlsm doc linked to this project. 

```
  '1a) Create a ticker Index

        tickerIndex = 0

    '1b) Create three output arrays
    
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single
        
        Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
        For i = 0 To 11

            tickerVolumes(i) = 0
            
        Next i
     
    '2b) Loop over all the rows in the spreadsheet.
        
        For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        'If  Then
        
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
    
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                
        'End If
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        'If  Then
            
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
                '3d Increase the tickerIndex.
            
                        tickerIndex = tickerIndex + 1
    
            End If
    
        Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i 
``` 

## Project Summary

### Advantages of Refactoring
Refactoring code can make a macro run faster, sometimes significantly faster, by designing the code in a simpler and more elegant way. This is very important if you have a large data set that takes a long time to run, especially if it is a data set that may continue to get larger over time. 

### Disadvantages of Refactoring
Refactoring code takes time and can even be a job in itself. A person who is being paid to consult a business through performing data analysis must weigh the benefit of spending additional time refactoring their code before submitting the deliverable to their client, especially if they’re being paid by the hour. Executives at a large company that employs multiple, or even hundreds, of data scientists may need to build a business case for hiring more analysts to simply refactor code and make it run more efficiently for internal and external data requests. 

### Pros and Cons in relation to this project
Refactoring the original VBA script was an important exercise in service to learning how to rethink code and make it run more efficiently. Outside of getting more experience with VBA coding and the concept of refactoring, it wasn’t critically necessary, because our data set was relatively small. If our goal is to get the code written as quickly as possible, then we would likely stick with the original macro. That being said, if we needed to run an analysis on all publicly-traded stocks, not just those for green energy companies, the time we spent refactoring the code would save us more time than it takes to refactor the code, which would be a value add. 
