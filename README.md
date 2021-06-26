# VBA of Wall Street

## Overview of Project
### Background
In order to help Steve with the analyzing wall street, I had created a workbook. Steve Loves the workbook created for him.
Now, to help his parents in understanding the stock market he wants to expand the dataset to include the entire stock market over last few years. Also he wanted to make it more efficient in order to get the quick system output.
 
### Purpose

Main purpose is to refactor the code. 
Sometimes code may work for small amount of data efficiently but may take long time to work for large amount of data. So to make the the process more efficient we can refractor the code. Here, Steve also wanted to check the time to process the code.

In this part we are going to deliver'
1. This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time
2. Detailed written analysis of the result.
 

## Analysis

So after analizing the output data for years 2017 and 2018 it looks like only two tickers ENPH and RUN for which the total daily volume has been increased and gives positive results. 
Here is the output of All Stock Analysis worksheet 

For Year 2017
[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/VBA_Challenge_2017-output.png)

For Year 2018
[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/VBA_Challenge_2018-output.png)


## Result

### Results Before Refactoring

Before refactoring the time to execute the code was 1.171875 seconds for 2017 and for 2018 it shows 1.160156 seconds.

For year 2017
[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/2017%20before%20refactoring.png)

For year 2018
[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/2018%20before%20refactoring.png)

### Results After Refactoring
There is the huge difference after refactoring the code it takes now 0.0859375 seconds for 2017 and for 2018 it shows 0.0703125 seconds.

VBA challange 2017
[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/VBA_Challenge_2017.png)

VBA Challange 2018
[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/VBA_Challenge_2018.png)

## Summary

1.What are the advantages or disadvantages of refactoring code?

Advantages

The Refactored code is well organized and easy to understand, so if there are any bugs or fixes in code it saves time to find the bug or fixes.
Code refactoring removes redundant or duplicate code and improve effectivness of the code.
It boosts the performance of the application or system. 

Disadvantages

Code refactoring takes too much time. If the code is complex there are high chances to lost in between somewhere.


2.How do these pros and cons apply to refactoring the original VBA script?

Original Code

Original code contains nested for loops, which increase the number of itrations hence it takes more time to process the code.
Here is the link for Original Code

[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/Original%20Code.png)

Refactored Code

Refactored code has only single for loop in order to process data more efficiently. Which also helps in reducing time to process the code and generate output.
Here is the link for Refactored code

[GitHub](https://github.com/rachanamule/stock-analysis/blob/4bddaca128455020f1ec8e88f68836e0385a88e8/resources/Refactored%20Code.png)

In this, the code was of only few lines so refactoring was helpful in such cases. 
