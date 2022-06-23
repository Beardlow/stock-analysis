# Stock-Analysis
## Analysis of Module 2 Stock Data

### Project Overview

This project was given to us by a financial advisor colleague who was in need of a faster way to analyze the trading volumes and the stock performance, of a list of stocks, for a given year. While using pivot tables, could have accomplished the same obejective, the use of VBA scripting and automation has sped the process up significantly. This same VBA code will work for future years as well, provided that the data is in the same format, and the same stocks are being analyzed.

### Results

#### 2017 Results
This anlaysis shows that 2017 was a pretty good year for the sampled market as a whole; however, the execption to this is the stock with ticker "TERP". TERP had a 7.2% decrease in the stock price from the beginning of the year 2017 to the end of the year 2017. Every other stock listed had a positive return for the year 2017.

![2017_Returns](https://github.com/Beardlow/stock-analysis/blob/main/2017_Returns.png)

#### 2018 Results
The Analysis fo the year 2018 tells a very different story. Only two stocks showed a positive return for the year 2018. These stocks are shown as tickers "ENPH" and "RUN". This may be indicative of a broader market event or slowdown; however, 2018 had roughly 139 million more trades occur amongst the stock listed than did 2017. A more detailed look at 2018 and the market environmental factors for the year may be needed. It could also be inferred that the tickers ENPH and RUN are more resilient to market downturns than the other stocks analyzed.

![2018 Returns](https://github.com/Beardlow/stock-analysis/blob/main/2018_Returns.png)

#### Performance of Results
The comperter processing time needed to run the results for both 2017 and 2018 were very similar or exactly the same. This is most likely due to the amount of data for both 2017 and 2018 being the same. See performance results for both year's analyses below.

##### 2017 Performance
![2017 Analysis Performance](https://github.com/Beardlow/stock-analysis/blob/main/VBA_Challenge_2017.png)

##### 2018 Performance
![2018 Analysis Performance](https://github.com/Beardlow/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Summary

#### What are the advantages or disadvantages of refactoring code?

#### Advantages

One of the advantages of refactoring code is that using existing code can provide you with a framework to work with. The framework given may also be properly working code, that just needs to be more efficient, so the results of the initial code could be used a a refernce to verify the correctness of the analysis. This also provides a benchmark for measuring performance of the script itself.

#### Disadvantages

A disadvantage to refactoring code is that you are sometimes working with code written by someone else. This means that you may have a difficult time understanding how the original code author was thinking at the time of creation. Another disadvantage is that the previous author may not have been diligent with the creation of code comments.  This can lead to confusion and frustration when trying to figure out what a particular piece of code is doing. Due to these reasons, sometimes, it may be necessary to start fresh and just use the original code as a reference.
