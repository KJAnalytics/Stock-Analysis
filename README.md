#  ***Stock-Analysis:***

##  Project Purpose 
Analysis of a range of stock returns to include code development in VBA, refactoring, and reductions in execution time.  As a result of the analysis, provide feedback to a client on a particular stock's performance with recommendations on additional stocks to consider.

##  Project Overview 
Data provided is a series of "green based" stocks with daily trading activity organized in tabs by year. Data is provided in exel, thus VBA was chosen to complete the analysis.
## Objectives and analysis Findings 
### Objectives:            

Refactoring code to complete analysis of stock performance including Total Transaction Volume and Return for a series of 12 stocks including comparison of code run times with the objective of reducing overall runtime. In order to accomplish this goal, coding activities were structured and build on basics to accomplish the Challenge objective.  This foundation included:           
-Evaluation of a single stock with ticker DQ for 2018, using daily stock volumes to calculate the total transaction volume and return on investment for the year.  This stock was chosen based on current clinet investment portfolio focusing on green energy stocks.
-  Comparison of performance 2017 vs 2018 for stock DQ.
-  Evaluation of an additional 12 stocks with analysis of performance over the same 2 year period.
-  Refactoring code to complete analysis including comparison of code run times with the objective of reducing overall runtime.
 ### Analysis takeaways  
-  Stock DQ chosen by the client's father experienced a 62% decline in return for 2018 after experiencing record returns in 2017 @ 199%.
-  After analysis of additional stocks, 10 of the 12 stocks evaluating experience declines in 2018 with DQ experiencing the largest losses.
-  Stocks ENPH and RUN are the faonly two stocks in the analysis set with positive returns year to year.  
-  Before making a "Sell" or "Buy" recommendation, additional years should be added to the analysis to betterunderstand the impact of trading volume on the performance of stock DQ as return decreased as trading volume trippled.

*******insert picture here

 ### Refactoring takeaways
 
Improvement of code to reduce overall run time has multiple approaches and may involve a series of steps to acheive fastest times.  In the refactoring, a reduction of 8.9% (0.1015 seconds) was accomplished using all the resources listed below.  Initial code run time for both years was 1.136719 seconds and finished with 1.035156 seconds refactored. Autocalculation is a time hog and with the small shifts seen in this program, imagine the impacts when working with thousands of lines of code.:
- Inactivated screen updates, status bar updates. animations, and Events.
-  Utilized With...End calls for multiple formatting calls for ranges.

*****insert examples here

Finally refactoring has pros and cons.
                        -  I found that my biggest issue copying text from other projects and using in a new project takes dilgence and planning.  Without proper planning simple issues like tag names, For/End loops, and not indenting code properly can lead to loads of time lost chasing syntax, overrun, and other programming errors.
                        -  To mitigate these risks - proper planning and sketching out code and tags in advance can reduce hours of time spent searching for an error that is only 1 letter different.
                        -  Finally - learn to walk away when the code isn't flowing.  The break can clear the head and lead to a solution.
                        -  Google, stackoverflow and youtube are my new best friends outside of class.
                        
*****insert pictures here.

### Results

-Starting from base green_stocks code and the prompts provided for the VBA_Challenge exercise, refactoring and validaton of code integrity was conducted.  Code initiated from a simple "Hellow World" in class and progressed through a series of mini projects each building on concepts learned in previous exercises. Concepts demonstrated include:
            - Simple text message to a worksheet in a Msgbox
            - Formatting work sheet titles, column headers, and color coded cells based on analysis results such as green/red for return performance for the stock analysis and generation of a colored checkerboard using a calculation of a Mod value to trigger Odd/Even status to set cell color.  See excepts from code below. attachments 1-3.
            - Class practice involving For/Next Loops and Nexted loops.
            -  Analysis of a single stock DQ to determine 2018 Retun information including a pushbutton to reset the sheet and a Message box to input the year for analysis.  Attachments 4-6
            -  All of the above rolled into the AllstocksAnalysisRefactored VBA Module and the reduction in code processing times.
           

****** insert pictures here
            
            
            
#### Notes below include the details of each commit for reference

####
1st commit uploaded green_stocks-Module 2 .xlsm adding 1st module Macro "Hello World" in text box on screen.  Also added code to call DQAnalysis sheet active and call oufor i from 2 to 3013.  Created title and column headders using 2 methods. 

#### 2nd Commit - uploaded green_stocks-Module 2-Rev1.xlsm File Also added code to call DQAnalysis sheet active and call out title and column headders using 2 methods.  First used Range function = "DAQO (Ticker: DQ)".  Second Cells(1, 1).Value = "DAQO(Ticker: DQ)". 

#### 3rd commit - begins data analysis using sheet for 2018.  Initial code pulls the column titles from worksheet "2018".  Next a For loop is created to cenerate the total volume of transactions when the stock is DQ.  Used 2 different method for defining the rows for analysis.  First hard coded rows 2 to 2013 from sheet 2018 which limits analysis going forward.  Second method still hard coded but used variables rowstart and rowEnd with set values providing limited flexibility but data additions require updates to that script.  Finally a flexible method was used using range and end calls where it will automatically update and include newest rows of data as only the starting point value (2 - for row 2 where data starts is coded) and this was demonstrated pulling data from sheet 2017. ![image](https://user-images.githubusercontent.com/106294465/172052102-27aeeac1-1f9e-4362-863b-3a8e4d80e5b6.png) Finally, the total volume of transactions for stock DQ were calculated for both sheets 2018 and 2017 and loaded into DQAnalysis sheet.  This was done using a For statement. Within the iteration Transaction_Total_Volume was declared a starting value of 0.  then each cell is checked for DQ (If Cells(i, 1).Value = "DQ" Then)  and if DQ the valued is added Transaction_Total_Volume = Transaction_Total_Volume + Cells(i, 8).Value the IF statement is then ended and the next iteration is repeated.  Once complete, analysis was conducede on sheet 2017 as well and dropped into the DQAnalysis sheet.  Critical learning  was when creating a message box to initially check the value, it was loaded into the iterative loop which made the message box open for each row.  This made the box     pop for every row.  Had to kill excel in order to stop the process and edit code to move out of the loop

#### 4th Commit - adds analysis for calculating return on the stock DQ.  This added 2 new variables startingprice and endingprice that were declared as double since the values were not integers and a calcultion is being completed that may result in a fractional result.  These variable are then used to calculate a return on a stock.  This follows in the loops used for the  total transaction volume to iclude the new calculation:  Cells(4, 3).Value = (endingprice / startingprice) - 1.  In the iterative loop starting price  and ending price are compared to the rows above and below the row that is currently being calculated. It is also refined to be looking at only DQ stock.  Analysis is repeated in a second loop for year 2017 as well. The comparison code is shown: ![image](https://user-images.githubusercontent.com/106294465/172054975-f6f118c1-47a9-4cfa-8a11-f38728ba10e3.png).  In the end, the spreed sheet now appears as below: ![image](https://user-images.githubusercontent.com/106294465/172055351-4cdf9e94-3429-467e-997d-9a246322e1ed.png)

#### 5th commit - covers 2 sets of activities
- class practice activity with nested loops generating data within rows and columns using standard set # of rows and columns that are predefineda second iteration with variable rows and columns. Once the cells are filled, the 3rd step was to overwrite that data with a sum of row and column numbers.  The final request was to clear the sheet.
- This was a fun activity and I found a neat way to use a random number generator to set the variables for the number of rows and columns.  In addition, I added a call at the top of the program to clear the worksheet of past values so that it can be run again from a clean sheet.![image](https://user-images.githubusercontent.com/106294465/172086375-d1210b38-9f37-4001-a401-25d809444042.png)
- The second half of the upload is the Stock Analysis sheet and module 3 code.  This activity focused on analysis of all stocks from 2018 and calculating Total Daily Volume and return on investment. Attached is the output of the code.![image](https://user-images.githubusercontent.com/106294465/172086966-403e0a18-1342-4b24-a04f-bbc53674082b.png)
- Key features include the nested loops to iterate through total daily volume and return calculations for an array of 12 unique stocks.  In the code, I researched stack overflow to determine how to write the Return data out as percentages.![image](https://user-images.githubusercontent.com/106294465/172088146-2714e308-ff4c-495e-8ce4-c86a19b33554.png)

-

            
