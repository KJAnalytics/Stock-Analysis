#  ***Stock-Analysis:***

##  Project Purpose 
Analysis of a range of stock returns to include code development in VBA, refactoring, and reductions in execution time.  As a result of the analysis, provide feedback to a client on a particular stock's performance with recommendations on additional stocks to consider.

##  Project Overview 
Data provided is a series of "green based" stocks with daily trading activity organized in tabs by year. As the data was provided in exel format, VBA was chosen to complete the analysis.
- Objectives for analysis included:
            -  Evaluation of a single stock with ticker DQ for 2018, using daily stock volumes to calculate the total transaction volume and return on investment for the year.
            -  Comparison of performance 2017 vs 2018
            -  Evaluation of an additional 12 stocks using analyze performance over the 2 years of data.
            -  Refactoring code to complete analysis including comparison of code run times with the objective of reducing runtime.
 
- Key takeaways:
- Stock DQ chosen by the client's father experienced a 62% decline in return for 2018 after experiencing record returns in 2017 @ 199%.
-  After analysis of additional stocks, 10 of the 12 stocks evaluating experience declines in 2018 with DQ experiencing the largest losses.
-  Stocks ENPH and RUN are the only two stocks in the analysis set with positive returns year to year.  
-  Before making a "Sell" or "Buy" recommendation, additional years should be added to the analysis to understand the impact of trading volume on the performance of stock DQ as return decreased at trading volume trippled.

Challenges/Learnings for this project:
-

##  
## 1st commit uploaded green_stocks-Module 2 .xlsm adding 1st module Macro "Hello World" in text box on screen.  Also added code to call DQAnalysis sheet active and call oufor i from 2 to 3013.  Created title and column headders using 2 methods. 
## 2nd Commit - uploaded green_stocks-Module 2-Rev1.xlsm File Also added code to call DQAnalysis sheet active and call out title and column headders using 2 methods.  First used Range function = "DAQO (Ticker: DQ)".  Second Cells(1, 1).Value = "DAQO(Ticker: DQ)".  
## 3rd commit - begins data analysis using sheet for 2018.  Initial code pulls the column titles from worksheet "2018".  Next a For loop is created to cenerate the total volume of transactions when the stock is DQ.  Used 2 different method for defining the rows for analysis.  First hard coded rows 2 to 2013 from sheet 2018 which limits analysis going forward.  Second method still hard coded but used variables rowstart and rowEnd with set values providing limited flexibility but data additions require updates to that script.  Finally a flexible method was used using range and end calls where it will automatically update and include newest rows of data as only the starting point value (2 - for row 2 where data starts is coded) and this was demonstrated pulling data from sheet 2017. ![image](https://user-images.githubusercontent.com/106294465/172052102-27aeeac1-1f9e-4362-863b-3a8e4d80e5b6.png) Finally, the total volume of transactions for stock DQ were calculated for both sheets 2018 and 2017 and loaded into DQAnalysis sheet.  This was done using a For statement. Within the iteration Transaction_Total_Volume was declared a starting value of 0.  then each cell is checked for DQ (If Cells(i, 1).Value = "DQ" Then)  and if DQ the valued is added Transaction_Total_Volume = Transaction_Total_Volume + Cells(i, 8).Value the IF statement is then ended and the next iteration is repeated.  Once complete, analysis was conducede on sheet 2017 as well and dropped into the DQAnalysis sheet.  Critical learning  was when creating a message box to initially check the value, it was loaded into the iterative loop which made the message box open for each row.  This made the box     pop for every row.  Had to kill excel in order to stop the process and edit code to move out of the loop
## 4th Commit - adds analysis for calculating return on the stock DQ.  This added 2 new variables startingprice and endingprice that were declared as double since the values were not integers and a calcultion is being completed that may result in a fractional result.  These variable are then used to calculate a return on a stock.  This follows in the loops used for the  total transaction volume to iclude the new calculation:  Cells(4, 3).Value = (endingprice / startingprice) - 1.  In the iterative loop starting price  and ending price are compared to the rows above and below the row that is currently being calculated. It is also refined to be looking at only DQ stock.  Analysis is repeated in a second loop for year 2017 as well. The comparison code is shown: ![image](https://user-images.githubusercontent.com/106294465/172054975-f6f118c1-47a9-4cfa-8a11-f38728ba10e3.png).  In the end, the spreed sheet now appears as below: ![image](https://user-images.githubusercontent.com/106294465/172055351-4cdf9e94-3429-467e-997d-9a246322e1ed.png)
## 5th commit - covers 2 sets of activities
- class practice activity with nested loops generating data within rows and columns using standard set # of rows and columns that are predefineda second iteration with variable rows and columns. Once the cells are filled, the 3rd step was to overwrite that data with a sum of row and column numbers.  The final request was to clear the sheet.
-- This was a fun activity and I found a neat way to use a random number generator to set the variables for the number of rows and columns.  In addition, I added a call at the top of the program to clear the worksheet of past values so that it can be run again from a clean sheet.![image](https://user-images.githubusercontent.com/106294465/172086375-d1210b38-9f37-4001-a401-25d809444042.png)
- The second half of the upload is the Stock Analysis sheet and module 3 code.  This activity focused on analysis of all stocks from 2018 and calculating Total Daily Volume and return on investment. Attached is the output of the code.![image](https://user-images.githubusercontent.com/106294465/172086966-403e0a18-1342-4b24-a04f-bbc53674082b.png)
-  Key features include the nested loops to iterate through total daily volume and return calculations for an array of 12 unique stocks.  In the code, I researched stack overflow to determine how to write the Return data out as percentages.![image](https://user-images.githubusercontent.com/106294465/172088146-2714e308-ff4c-495e-8ce4-c86a19b33554.png)

-

            
