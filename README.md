# Green Stock Performance Analysis

## Purpose

#### Initial approach

We were given the task of completing a comparative analysis of the performance of green stocks between the years 2017 and 2018 based on the dataset with which we were provided - two Excel worksheets, each containing tabulated data for the corresponding year. Each spreadsheet contained tabulated values corresponding to twelve different stock tickers, showing each respective stock's performance during a series of days. For each day (each row), we had information about the given stock at its highest, its lowest, at opening and closing, and the volume that was traded on that day. 

#### Programmatic solutions

We performed analysis using Macros in VBA to create code that compiled and organized information from the dataset. In order to gauge each stocks' performances, the programmatic approach helped us dive into our dataset in quicker and more efficient ways. 

In a new worksheet, we synthesized information about each stock, creating a table with three columns: one identifying the ticker (stock), one that showed the total daily volume traded by each stock during the course of the year, and one that summarized the percentage of return of that stock, giving a sense of profitability of each investment. A positive returned showed green, and a negative returned showed red, thereby producing a visual effect that helped in comparing and contrasting the stocks' performances on each year. 

We created an interactive button that initiates the analysis, with a pop-up message box that prompts the selection of the year in question: 2017 or 2018 in this case. The year that was selected then shows on the first cell at the top left hand corner, and the analysis corresponding to that year is displayed in the table below.

Transcending the use of basic Excel formulas and applying the use of VBA programming, we were able to create interactive tools that, optimized even further by refactoring code, are capable of populating tables with relevant information; making our understanding of data analysis fast, captivating and dynamic. 

## Results

#### Comparing Stock Performances in 2017 and 2018

At first glance, what we saw is that returns for the totality of the stocks showed more negative results in 2018 than in 2017. Since we used colors to distinguish the cells with positive and negative returns, 2017 appeared mostly green, and 2018 appeared mostly red. We can infer from this that the general state of the market was distinct on each of these years.

Looking closely at the performance of each stock, we observed that the TERP ticker was the only ticker that showed consistently negative results in 2017 as well as in 2018; which might be indicative that, given the overall decrease in returns in 2018, and a generally less conducive stock market on that year, the TERP ticker would still have been on its own path, and potentially not a good investment target altogether, considering the data we have. The ENPH and RUN tickers, in turn, tell interesting individual stories, and we decided  to look at these two stocks more closely. 

In 2018, the **ENPH** ticker showed a decrease in numbers for its returns, but its returns remained positive. Hence, ENPH we may infer from this data that ENPH stocks might be a strong contender, as maintaining a stable positive standing in spite of general instability and oscillation in the market can be a reassuring factor.

The **RUN** ticker, in turn, actually showed significant increase in returns from 2017 to 2018, against the general trend. This might signify, in a preliminary analysis, using the data we have, that RUN would have been the most profitable green stock to have invested in between the years of 2017 and 2018.

These screenshots illustrate the comparison of performances of green stocks in the years 2017 and 2018:




![This is an image](VBA_Challenge_2017_stock_analysis.png)





![This is an image](VBA_Challenge_2018_stock_analysis.png)





#### Comparing Programmatic Performances

In order to automate our analysis, we created **VBA Macros** which ran effectively, outputting data into the new worksheet, populating the new spreadsheet we had created,  generating comparisons between stocks in a clear way, and allowing us to perform a fair analysis of the stocks.

That said, after creating the VBA Macros, we observed that even better performances could be achieved by the VBA codes by **refactoring** the original ones that we had created.



#### Original Code Explained

In the **original code**, we iterated through each of the the rows using a *for loop* to find the total daily volume for each stock, the starting and ending price. Within this *loop* we started the total volume to *zero* when we established its variable; and  started another *for* *loop* still within it, that iterated using another variable to find the total volume, starting and ending prices using *if statements*.  As a final argument of the first loop we outputted the total daily volume and the return - using starting and ending price - for every row, as seen here:

```
'Loop through the tickers
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through rows in the data
        Sheets(yearValue).Activate
        
            For j = 2 To RowCount
        
                'Find the total volume for the current ticker
                If Cells(j, 1).Value = ticker Then
                
                    totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
                
                'Find the starting price for the current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                    startingPrice = Cells(j, 6).Value
                
                End If
                
                'Find the ending price for the current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                    endingPrice = Cells(j, 6).Value
                
                End If
                
            Next j
    
    'Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
```



#### Refactored Code Explained

In the **refactored code**, we started by creating a new Ticker Index variable. We then created three output arrays, for Volumes, Starting Prices and Ending Prices. We then created a *for loop* that initialized all ticker volumes to *zero* (0). After that we used a separate *for loop* to loop through the rows iterating the ticker volumes for each ticker, and within it also we used  *if statement*s to identify the starting and ending prices. We then started a new *for loop* that looped through the arrays we had created initially to output the total daily volumes for each stock, and the return, based on outputs of Starting and Ending Price arrays, as seen here:

```
1a) Create ticker Index
    tickerIndex = 0
   
     '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
  
    
    '2b) Loop over all the rows in the spreadsheet
    For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                        tickerStartingPrices(tickerIndex) = Cells(i, 3).Value
            End If
            
            '3c) Check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                        
            '3d) Increase the tickerIndex
            tickerIndex = tickerIndex + 1
            
            End If
            
    Next i
    
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
         
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i
    
```



#### Comparison of Execution Times

By verifying the execution times of the original and refactored codes we could observe that the refactored code **performed better** and **ran faster** in the analysis of both 2017 and 2018 dataset.

Below are the screenshots that reveal this improvement. 

This is the execution time of the original code in the 2017 analysis:


![This is an image](VBA_Challenge_2017_excecution_original.png)

This is the execution time of the refactored code in the 2017 analysis:


![This is an image](VBA_Challenge_2017.png)

This is the execution time of the original code in the 2018 analysis:


![This is an image](VBA_Challenge_2018_excecution_original.png)

This is the execution time of the refactored code in the 2017 analysis:


![This is an image](VBA_Challenge_2018.png)



## Summary 

#### Questions

##### What are the advantages or disadvantages of refactoring code?

Advantages of refactoring code:

- Refactoring a code can enhance the performance and execution of code in VBA Macros, thereby rendering a faster analysis.
- Code can be reutilized and can appear cleaner and more accessible for future use by yourself any member of your team.
- Refactored code can make more viable its improvement and the utilization of its various aspects in different applications.

Disadvantages of refactoring code:

- Refactoring can be time consuming, especially when the process isn't seamless. Depending on how disorganized the original code would have been, attempting to refactor it can lead to misreading and confusion.
- The programmers can get side-tracked, and a project that had been working well may get delayed due to errors generated in the process.
- Potential complexities encountered in reutilization of code may result in verification that the advantages of refactoring might not always outweigh the cost of time and energy reformulating code; and in some cases just creating new code altogether may be the most effective approach.

##### How do these pros and cons apply to refactoring this original VBA script?

- Pros of refactoring this script:
  - The pros of refactoring this script were pretty self-evident. The script we originally created for this analysis was effective. The refactored code, however, was more elegant, more intelligible, more accessible, and its execution was faster in both its analysis. 
  - As a tool for data analysis, using the faster-running refactored code can be advantageous in that it can provide opportunities for more efficient further analysis of more dataset, including more years, more tickers and a wider scope of data. 
  - If another individual or team would take on this project, they would likely have an easier time understanding the refactored code, and may be able to add different arrays and other features to the code.

- Cons of refactoring this script:
  - In the process of refactoring code, errors can be made and results may be misinterpreted and then used incorrectly throughout the analysis.
  - Upon expansion and reformulation, the process of refactoring and reutilizing refactored code can be more time consuming than the effort it would take to create a new code altogether for a different application.

#### Conclusion

In conclusion, I believe that the use of refactored code in this particular analysis was beneficial and a worthwhile effort. Refactoring our original code was not only advantageous because of the results - the refactored code's better and more efficient execution time and performance. Those benefits speak to how best we could present and execute our analysis. Its also worth noting that we found, in the process of refactoring, that we were able to better understand how to manipulate and organize the information we had from our dataset; and then organize our thought process through grouping the information we wanted to extract, and were able to retrieve it into our output table more neatly and succinctly.
