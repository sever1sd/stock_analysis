# Stock Analysis with VBA 
## Overview

**Purpose**

This analysis was performed in order to evaluate the stock market performance of select "green" companies. VBA code was used to collect and aggregate stock data for the identified companies. Performance results per stock were calculated and printed on a new sheet within the workbook. The main goal of this project was to evaluate and refactor code. The orginal code was provided in a .vbs file and refactoring was required in order to increase efficiencies.

**Summary of VBA Code**

The script was designed to collect and aggregate information on each stock ticker provided. The code allows for the user to select which year (2017 or 2018) of data to evaluate. The script then loops through the data for the appropriate year, aggregating the total volume by stock ticker. It also finds the first and last instance of a stock ticker and its associated stock price. These were then used to calculate the return rate for the given year. Finally, the results were outputted onto a separate tab and formatted for easy-to-read analysis.

The workbook in which this analysis was performed can be downloaded from [this link](Challenge/VBA_Challenge.xlsm). The code used to run this macro is available in this workbook under "Module 2."

## Results

The original code was provided via a .vbs file. A new module was created within the workbook (named Module 2) and the original code was copied over to the module. The code contained the framework for setting variables and outputting results, but did not have the required loops to perform the analysis.

### VBA Code Used

**Setting Variables**

The first part of the code was dedicated to setting the variables, setting the stage for the output file, and creating an array for the tickers variable. This array was used to cycle through the tickers via a loop. Additionally, a user-input message box was set as a variable in order for the end results to be customizable to a year. This user-input was defined as a variable and was used to identify the correct worksheet to run the analysis on and update the header for the respective year.

![alt text](https://github.com/sever1sd/stock_analysis/blob/5d38201dbd908dda525fc394eb194cd5cef9a8bd/Challenge/Resources/Code%20sets%20variable.png)

**Preparing the loop**

The refactoring of code began at this point of the project. A new variable was set to identify the number of rows in later loops. The code for "rowCount" was provided, but a new variable needed to be set in order for this code to function. "tickerIndex" was created as a variable and set to zero, as well as 3 arrays which would be used to identify and aggregate data in the main loop.

![alt text](https://github.com/sever1sd/stock_analysis/blob/5d38201dbd908dda525fc394eb194cd5cef9a8bd/Challenge/Resources/Code%20preps%20loop.png)

**Setting the Ticker Index to 0 and Main Loop**

A for loop was created to set the variable "tickerVolume" to zero each time the script looped through the data. This was necessary to prevent unintentional stacking of volumes unrelated to the active ticker in the loop. Then a for loop was established which does most of the heavy lifting for the code. This loop will cycle through each ticker via the "tickerIndex" array and aggregate volume per ticker and establish the starting and ending stock price for the given year. The variable "tickerVolumes" increases the volume of the given ticker, which is established in the loop using the "tickerIndex" array. A series of conditional statements were used to set the starting and ending price of a ticker as well as to increase the "tickerIndex" array by 1 if the loop had cycled through all possible options for the active stock ticker.

![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/Code%20volumes%20Reset%20and%20Main%20loopv1.4.png)

**Output**

A for loop was used to write the results of the main loop onto a sheet within the workbook. The loop fills a cell with the appropriate data based on the tickerIndex array. A calculation is also performed within the loop to identify the rate of return for each ticker. 

![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/Code%20Output%20loop.png)

**Formatting**

Finally, the script reformats the data to be easier consumption and analysis. A loop is used to apply conditional formatting to the rate of return column to easily identify positive and negative values. This piece of the code was not a part of the refactoring.

![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/Code%20Formatting.png)

**Final Charts**

The script results in the following output charts:

  <ins>2017<ins>
  
![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/2017%20Stock%20Analysis%20Results.png)

  <ins>2018<ins>
  
![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/2018%20Stock%20Analysis%20Results.png)

## Summary
**Advantages of Refactoring Code**
  
 It is clear that refactoring code is a key part of the use and creation of scripts. The initial code built for this exercise throughout the course, while functional was substantially slower than the refactored code. 
    
 ***The original script timer output:***

  <ins>2017<ins>

![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/2017%20Original%20Stock%20Analysis%20Script.png)
  
  <ins>2018<ins>
  
![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/2018%20Original%20Stock%20Analysis%20Script.png)

 ***The refactored script timer output:***

  <ins>2017<ins>

![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/2017%20Refactored%20Stock%20Analysis%20Script.png)
  
  <ins>2018<ins>
  
![alt text](https://github.com/sever1sd/stock_analysis/blob/e6293ea5b9be4f86a79448f653b80bca8a7dd834/Challenge/Resources/2018%20Refactored%20Stock%20Analysis%20Script.png)
    
 The difference in time to run is fairly substantial. The refactoring of the code reduced the amount of time and resources the script used to run the analysis. While it was only a few seconds in this macro, for larger datasets or analysis the additional speed and lower use of computer memory could make a huge impact.
    
**Pros and Cons of Refactoring original VBA Script**
