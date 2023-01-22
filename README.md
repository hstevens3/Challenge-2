# Challenge 2 Refactoring VBA

# Refactor VBA Code for Stock Analysis

## Overview of Project

Using the starter code provided, this challenge is to refactor the VBA code from Module 2. The objective is to learn how to use arrays and compare the execution time for each approach to see which is faster. 

## Results

### Analysis of Stock Performance

Overall, stocks performed significantly better in 2017 than 2018. DQ had the highest return (199.4%) of any stock in 2017. TERP was the only stock in 2017 that had a loss. In 2018, only two stocks ENPH and RUN had a positive return at 81.9% and 84.0% respectively.

### Analysis of Execution Time

The original code uses nested for loops and reads the data rows for each iteration of the inner loop. The refactored code uses arrays and loops through the data rows only once. The refactored code runs in approximately half the time of the orginal code. If there was a larger data sample with more tickers, the time savings would likely be even greater. The execution time for the runs for each year are shown below.

---

![Refactored Code Execution Time for 2017 ](/Resources/VBA_Challenge_2017.png)

---

![Refactored Code Execution Time for 2018](/Resources/VBA_Challenge_2018.png)

## Summary

- Refactoring code takes additional time, but in cases where the code will be used repeatedly it pays off. Time will be recouped as the code is repeatedly used. When refactoring, the code may be slimmed down and made easier to read and maintain. 
---
- Of course, refactoring should be done on a backup copy and only promote/replace the original code once the refactored code has been tested.
---
- In the stock analysis case, the refactored code is easy to read and can be extended for additional stock tickers without many additional updates since the index variable for the array is used.