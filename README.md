# Stock-Analysis
## Analysis of Module 2 Stock Data

### Project Overview

This project was given to us by a financial advisor colleague who was in need of a faster way to analyze the trading volumes and the stock performance, of a list of stocks, for a given year. While using pivot tables, could have accomplished the same obejective, the use of VBA scripting and automation has sped the process up significantly. This same VBA code will work for future years as well, provided that the data is in the same format, and the same stocks are being analyzed.

### Results

#### Code Used for Analysis

Some important metrics are captured for each of the listed stocks. These metrics are the total vlome traded for a given year and and performance for a given year. The code used to calculate these metrics is below.

##### Total Volume Code

```
If Cells(i, 1).Value = tickers(tickerindex) Then
    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
End If
```

##### Total Retun Code

```
If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
    tickerStartingPrices(tickerindex) = Cells(i, 6).Value
End If
                    
'3c) check if the current row is the last row with the selected ticker
    'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
          If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
              tickerEndingPrices(tickerindex) = Cells(i, 6).Value
          End If
                
'3d Increase the tickerIndex
                  
          If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerindex = (tickerindex + 1)
          End If
```
The above code determines the starting and ending prices for a given year for a given ticker.  These values are used in the following code to determine the overall return for the year.

```
Cells(4 + j, 3).Value = (tickerEndingPrices(j) / tickerStartingPrices(j)) - 1
```

##### 2017 Results
This anlaysis shows that 2017 was a pretty good year for the sampled market as a whole; however, the execption to this is the stock with ticker "TERP". TERP had a 7.2% decrease in the stock price from the beginning of the year 2017 to the end of the year 2017. Every other stock listed had a positive return for the year 2017.

![2017_Returns](https://github.com/Beardlow/stock-analysis/blob/main/2017_Returns.png)

##### 2018 Results
The analysis for the year 2018 tells a very different story. Only two stocks showed a positive return for the year 2018. These stocks are shown as tickers "ENPH" and "RUN". This may be indicative of a broader market event or slowdown; however, 2018 had roughly 139 million more trades occur amongst the stock listed than did 2017. A more detailed look at 2018 and the market environmental factors for the year may be needed. It could also be inferred that the tickers ENPH and RUN are more resilient to market downturns than the other stocks analyzed depending on what is found through additional research.

![2018 Returns](https://github.com/Beardlow/stock-analysis/blob/main/2018_Returns.png)

#### Performance of Results
The computer processing time needed to run the results for both 2017 and 2018 were very similar or exactly the same. This is most likely due to the amount of data for both 2017 and 2018 being the same. See performance results for both year's analyses below.

##### 2017 Performance
![2017 Analysis Performance](https://github.com/Beardlow/stock-analysis/blob/main/VBA_Challenge_2017.png)

##### 2018 Performance
![2018 Analysis Performance](https://github.com/Beardlow/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Summary

#### What are the Advantages or Disadvantages of Refactoring Code?

#### Advantages

One of the advantages of refactoring code is that using existing code can provide you with a framework to work with. The framework given may also be properly working code, that just needs to be more efficient, so the results of the initial code could be used a a refernce to verify the correctness of the analysis. This also provides a benchmark for measuring performance of the script itself.

#### Disadvantages

A disadvantage to refactoring code is that you are sometimes working with code written by someone else. This means that you may have a difficult time understanding how the original code author was thinking at the time of creation. Another disadvantage is that the previous author may not have been diligent with the creation of code comments.  This can lead to confusion and frustration when trying to figure out what a particular piece of code is doing. Due to these reasons, sometimes, it may be necessary to start fresh and just use the original code as a reference.
