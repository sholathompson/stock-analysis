# Overview of Project 
## Background of Project
Our client, Steve is supporting his parents with investing their money into Green Energy options. Steve’s parents are intent on “investing all their money into DAQO” a new energy corporation that specializes in creating Silicon wafers for solar panels. Our client expressed that he would like to diversify his parent’s portfolio to include other forms of green energy and would like analysis on other forms of bioenergy to include hydroelectricity, wind energy, thermal energy and bioenergy. Client has provided an excel document with data for DAQO as well as 11 other green energy corporations to analyze.

## Purpose of Project
The purpose of this project was to analyze DAQO as well as a portfolio of 11 other green energy investment options. Analysis will focus on the total number of shares traded throughout the day (i.e. total daily volume) as well as the percentage difference in price from the beginning of the year to the end of the year (i.e. yearly return) for each stock. 
Of specific note, Steve may want to look at a different set of stocks in the future and has requested that the analysis include an easy to use, flexible macro for running multiple stocks. 

## Deliverables 
- DAQO 
  - Daily Volume
  - Yearly Return
- All Green Energy Stocks (N =12)
  - 2017 & 2018 Daily Volume/Yearly Return 
  - Automated buttons that will run the macro automatically for any year identified. 
- Message Box that pops up to tell how long it takes for the VBA code to compile the results. To accommodate this we refractured the VBA script.

# Results 

## DAQO stock analysis 

![Table 1- DAQO Daily Volume and Yearly Return.png](https://github.com/sholathompson/stock-analysis/blob/main/Table%201-%20DAQO%20Daily%20Volume%20and%20Yearly%20Return.png)

As Table 1, illustrates, DAQO yearly return dropped 63% indicating that this is a risky investment for Steve's parents at this time.

## All Stocks Analysis 

### 2017 and 2018 comparisons

#### 2017 All Stocks Daily Volume and Yearly Return
![VBA_Challenge_2017.png](https://github.com/sholathompson/stock-analysis/blob/d5c59cf644a0d793225bd696e0fe658a6312f62d/VBA_Challenge_2017.png)

#### 2018 All Stocks Daily Volume and Yearly Return

![VBA_Challenge_2018.png](https://github.com/sholathompson/stock-analysis/blob/d5c59cf644a0d793225bd696e0fe658a6312f62d/VBA_Challenge_2018.png) 

The above tables suggest that in 2017 Green Energy corporations investments were profitable indicated by yearly return data being positive for all companies except Terra Form Power                                                                                                                                                                                                                                    (ticker: TERP) 

In 2018, most of the green energy companies, indicated a negative yearly return on investment. The two companies that indicated a positive yearly return on investment were Enphase Energy (ticker: ENPH) and SunRun (ticker: RUN)

 ### Speed of VBA code:

Original speed of VBA code to compile the 2018 results
![Resources/how long the scripts take to run 2018.png](https://github.com/sholathompson/stock-analysis/blob/0b3c736a0545cacb8a28a7b86cc2c74394cb2024/Resources/how%20long%20the%20scripts%20take%20to%20run%202018.png)


Refractor speed of VBA code to compile the 2018 results
![Resources/how long the scripts take to run 2018 (Refactored).png]( https://github.com/sholathompson/stock-analysis/blob/0b3c736a0545cacb8a28a7b86cc2c74394cb2024/Resources/how%20long%20the%20scripts%20take%20to%20run%202018%20(Refactored).png)

As the above screen shots suggest, that the execution time for the refactored script is significantly faster than the execution time of the original script. 


## Summary
#### The advantages or disadvantages of refactoring code
Refactoring makes the code more efficient thus making programming faster, easier to grasp and easier to maintain. Refactoring makes coding is more succinct, less complex and easier for developers to read at a later date and make updates without changing the functionality of the software. Further, refactoring allows you to notice and fix coding errors that may go unnoticed in its original script.

One disadvantage of refactoring is that it is possible to make a coding error that would take a lot of time to correct. Depending on the depth of the problem you may not know how to correct it, thus having to revert to the original script and restart.  

#### Pros and cons apply to refactoring the original VBA script

Refactoring made the VBA script clearer, cleaner and simpler to follow. However, if you’re not skilled, refactoring VBA you may create errors or change the functionality of the software. Refactoring VBA code can be quite time consuming for a beginner coder.
