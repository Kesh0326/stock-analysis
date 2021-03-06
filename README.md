# Stock-analysis

## Project Overview
Analyze a set of stocks based on their volumes and returns to determine which stock is worth investing in

## Analysis & Results

### Approach for determining Returns

#### Determination of ticker starting price 
```
If cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = cells(j, 6).Value
```
#### Determination of ticker ending price
```
If cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = cells(j, 6).Value
```
#### Determination of Returns
```
cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```
### Results

#### 2017 Stock Results

![2017 Results](https://raw.githubusercontent.com/Kesh0326/stock-analysis/master/2017%20Stock%20Results.png)

#### 2018 Stock Results

![2018 Results](https://raw.githubusercontent.com/Kesh0326/stock-analysis/master/2018_Stock_Results.png)

#### 2017 original script runtime results

![2017 original script runtime](https://raw.githubusercontent.com/Kesh0326/stock-analysis/master/2017_Original_Runtime.png)

#### 2017 refactored script runtime results

![2017 refactored script runtime](https://raw.githubusercontent.com/Kesh0326/stock-analysis/master/Resources/VBA_Challenge_2017.png)

#### 2018 original script runtime results

![2018 original script runtime](https://raw.githubusercontent.com/Kesh0326/stock-analysis/master/2018_Original_Runtime.png)

#### 2018 refactored script runtime results

![2018 refactored script runtime](https://raw.githubusercontent.com/Kesh0326/stock-analysis/master/Resources/VBA_Challenge_2018.png)

## Summary

       - Refactoring code enabled shorter runtimes, and more optimized code performance in terms of formatting and structuring. However a potential downside is the          addition of new bugs that can arise through human error thus impacting code performance.
       - As seen from the original and refactored runtime images above, it is evident that the runtime for the refactored code is about 4 times faster than the              original code. 

