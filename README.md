# stock-analysis
##VBA analysis
VBA Stock Analysis
 
Introduction

In this Module 2 challenge we are going to refactor the code to make the code more efficient. We will then check whether the refactoring has helped to decrease the run time of the VBA script.

Challenge Background

Steve just finished his finance degree and an excel savvy. His parents would like to invest in stocks and asked his expertise to take a decision. Here we have used Excel with VBA scripting to analyze the entire stock. We helped Steve to create a workbook “Analyze 2017 and 2018 Stock” and he is pretty much happy with that. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although the code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

Analysis

First download the challenge_starter_code.vbs and rename it to VBA_Challenge.vbs. Then create a Resource folder to hold the screenshots. Rename the green_stocks.xlsm to VBA_Challenge.xlsm then add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor. Then start refactoring the code. 
These are the changes. 
•	First add a ticker index
•	Then create three arrays named tickerVolumes with Long data type, tickerEndingprices with Single data type, and tickerStartingprices with Single data type as well
 
•	Create a for loop to initialize the tickerVolumes to zero and then create a for loop that will loop over all the rows in the spreadsheet
 
•	Loop over all the rows in the spreadsheet and finding totalvolume, ticker starting price and ticker ending price for each stock
 
 
•	Then increase the Tickerindex by 1
 

•	Use a for loop to loop through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in the spreadsheet
 
•	Finally, run the stock analysis, then confirm that the stock analysis outputs for 2017 and 2018 are the same as they were in the module
•	Time taken to run the code when we input 2017
 
Result for 2017
 
•	Time taken to run the code when we input 2018
 
Result for 2018
  
From this we can conclude that after refactoring the code it took less time to run the code and the results are same as the “All stock analysis”.
Advantages
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality, you just want to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
By refactoring our “All stock analysis” code, we were able to decrease the time for running the code.

Cons
Even if we are running the same code on the same spreadsheet for same year we are getting different running time.




