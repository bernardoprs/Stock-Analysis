# Stock Analysis 
## Purpose
This exercise aimed to edit the existing VBA code to optimize the running time and the efficiency of the code without taking anything from the analysis. The task is to analyze the performance of 12 stocks over two years to suggest where to invest. 
## Analysis
The data presents the ticker value of 12 different stocks over 2017 and 2018.  The data sets include the opening price, the highest price, the lowest price, closing time, and the total volume on a specific date.

 This exercise aims to analyze what the database has and quantify the annual trading volume and the yearly return. In 2017 the returns on the majority of the stocks were positives, with four stores having returns of +100%, only one of the stocks had a 7.2% negative outcome.  In 2018 the result was the complete opposite having most of the 12 stocks have negative returns for the year. Two stocks had positive returns on both years: “ENPH” and “RUN.” Based on the information, I would pay more attention to these two companies than the rest due to their extreme volatility. 

 Although both my original script and the refracted code ultimately did the same task, the refracted code is much more easily understandable and straightforward. 

 My original script took around .3 seconds to run the analysis, and the refracted code took an average of .1 seconds to run, cutting the time significantly. 


!<img width="257" alt="Screen Shot 2021-05-08 at 5 54 00 PM" src="https://user-images.githubusercontent.com/81016335/117554966-18229780-b029-11eb-9144-fe40f2ea2ba2.png">

<img width="261" alt="Screen Shot 2021-05-08 at 5 53 49 PM" src="https://user-images.githubusercontent.com/81016335/117554975-25d81d00-b029-11eb-8505-d5aadba69caf.png">

<img width="255" alt="Perfomance 2018" src="https://user-images.githubusercontent.com/81016335/117554978-2ffa1b80-b029-11eb-84d2-557c1fab6d49.png">

<img width="259" alt="Perfomance 2017 " src="https://user-images.githubusercontent.com/81016335/117554989-41432800-b029-11eb-9d40-099a6713f6da.png">

I’ve indicated parts of the code to indicate the difference from the two versions. 
Original Code:

  ```ruby   
     '3.Prepare for the analysis of tickers.
        '3 a)Initialize variables for the starting price and ending price.
            Dim startingPrice As Double
            Dim endingPrice As Double
        '3 b)Activate the data worksheet.
        Worksheets("2018").Activate
        'Find the number of rows to loop over.
            RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    '4.Loop through the tickers.
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
       '5b) get starting price for current ticker
       If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If
       '5c) get ending price for current ticker
       If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
```
   
  

Refracted Code
```ruby
   '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
  Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
            
            
       
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
           If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If
            

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
            
      
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
       
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i
 
```
## 	What are the advantages or disadvantages of refactoring code?
Much like proofreading an essay after writing it, refactoring code might provide a much more straightforward and precise way to come to the conclusion one might need to. Refactoring code has many advantages, but the most important, in my opinion, is efficiency. Creating a solution to a problem is not enough if that solution makes problems elsewhere. 
A disadvantage of refactoring is that if one spent too much time trying to optimize code, the original purpose of the document might be lost during an extended refactoring process while addressing peripheral goals.
## 	How do these pros and cons apply to refactoring the original VBA script?
The most significant advantage that one might have while refactoring the original script I had was to improve the script's efficiency. I manage to several milliseconds to the time that it took to process all of the information. Considering that this is a relatively small data set, it might seem like not improving performance. Still, if we are handling a much bigger data set, it might prove a very significant improvement. 
