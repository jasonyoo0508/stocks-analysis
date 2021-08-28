# stocks-analysis

## Background

Steve wants to analyze a group of 12 green stocks to support his parent’s investment decisions. In order to accomodate Steve's objective, we need annual volume and return on investment on the 12 stocks.

## Purpose

Steve wants to analyze a multiple stocks. This may increase the amount of time it takes the analysis to produce results and we need to improve the workbooks efficiency by refactoring the VBA code. 

## Result

3 new arrays were created: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price

The above 3 arrays store performance data for each stock when a for loop runs analysis on them.


## Findings

The execution time improved from average 1 seconds to approximately 0.25 seconds for both years. 
