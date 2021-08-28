# stocks-analysis

## Background

Steve wants to analyze a group of 12 green stocks to support his parent’s investment decisions. In order to accomodate Steve's objective, we need annual volume and return on investment on the 12 stocks.

## Purpose

Steve wants to analyze a multiple stocks. This may increase the amount of time it takes the analysis to produce results and we need to improve the workbooks efficiency by refactoring the VBA code. 

## Result

3 new arrays were created: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price

The above 3 arrays store performance data for each stock when a for loop runs analysis on them.


## General Analysis

![VBA_Challenge_2017](https://user-images.githubusercontent.com/89154507/131204885-847eaffb-44fb-47c1-87c2-e863ccdae972.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/89154507/131204886-e8c680b8-5926-4cff-a4ef-aa3afd53133e.png)


Only 2 of the 12 stocks, ENPH and RUN produced a positive investments in both years. Overall, there was a decrease in the market. 

## Findings

The execution time improved from average 1 seconds to approximately 0.25 seconds for both years. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/89154507/131204808-62aa64c1-8cf2-4266-a9d7-a7d43b213422.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/89154507/131204811-cc43abbd-a05d-4875-838e-7ccae106059c.png)

## Pros & Cons

Some of the obvious pros for the findings would be that the execution time drastically increased from average run time of 1 second to 0.25 seconds.

Some of the Cons could be that updating the code itself. If one step is messed up during the coding, the whole worksheet could be confusing and lost. 
