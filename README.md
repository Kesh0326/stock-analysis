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

![2017 Results]

#### 2018 Stock Results

![2018 Results]

#### 2017 original script runtime results

![2017 original script runtime]

#### 2017 refactored script runtime results

![2017 refactored script runtime]

#### 2018 original script runtime results

![2018 original script runtime]

#### 2018 refactored script runtime results

![2018 refactored script runtime]

## Summary

Refactoring code enabled shorter runtimes, and more optimized code performance in terms of formatting and structuring. However a potential downside is the addition of new bugs that can arise through human error thus impacting code performance

