# Analysis of Refactoring VBA code and Performance Measurement

## Overview
The purpose of this project was to improve the performance of existing VBA code. The objective of the original and refractured VBA code is to list the stock ticker, total daily volume, and return for each stock in a given data set. Both versions of the VBA code perform the following. 
1.	The user runs the macro 
2.	The user enters a year 
3.	The output is populated 
4.	A pop-up box informs the user of the number of seconds for which the code ran

## Results 

### Analysis

Both the original and refactored code create variables for the start and end time for the timer, create and initialize a variable to store the users year input, start the timer to start counting number seconds the code runs, set the value of cell A1 to a title for the output, set the values of cells C1 to C3 to be the header row for the output, initialize an array for the tickers, and gets the number of rows to loop through. 

![Start of Code]() 
Start_of_code

The original code utilizes nested for loops. The code in the outer loop starts by setting the value of the ticker variable to corresponding value in the ticker array. It also sets the initial value of the totalVolume variable to 0. 

![Og Start Outer Loop]()

The inner loop loops through each row of the data work sheet. First, the code inside this loop compares the value of the ticker array set in the outer loop to the value of the ticker cell in the data worksheet. If the values match, the totalVolume variable in increased by the value of the volume cell in the data worksheet. Second, a condition is used to determine and sets the value of the startingPrice variable. Third, another condition is used to determine and sets the value of the ending price. This concludes the inner loop. 

![Og inner Loop]()

After the inner loop the code continues in the outer loop to generate the output. This loop will repeat 12 times for each of the ticker symbols in the ticker array. 

![OG end of outer loop]()
Lastly, the endTime variable is set to timer to stop counting the number of seconds the code ran and a message box is set to display the number of seconds the code ran.   

![OG End Sub]()

The original code does not include text or conditional formatting. The formatting for this code is done in a separate macro and is not included in the count of seconds. 

![OG formatting]() 

The Refactored utilizes multiple arrays and separate for loops. After the ticker array and the rowCount are initialized, a variable is created for tickerIndex and three more arrays are created. 

![RF Var and Arrays]()

Next the refactored code creates a loop to initialize the value of the tickerVolume array. This creates 12 ticker volumes and sets the starting value for each to 0. 

![RF tickerVolumn Loop]()

The refactored code contains a second loop that loops through each row of the data worksheet to update values of the arrays. First, the tickerVolume variable in increased by the value of the volume cell in the data worksheet for the corresponding tickerIndex.  Second, a condition is used to determine and sets the value of the tickerStartingPrice variable. Third, another condition is used to determine and sets the value of the ending price and increase the tickerIndex after the row of the tickerEndingPrice. This concludes the second loop. 

![RF Loops]()

The refactored code contains a third loop that loops through each of the arrays to display their output. 

![RF outputs]()

Next, the refactored code contains the code for text and conditional formatting. Unlike the original code, the formatting is included in the same macro. 

Lastly,  the endTime variable is set to time to stop counting the number of seconds the code ran and a message box is set to display the runtime.
 
![RF end sub]()

### Improvements 
![OG 2017 Runtime]()
![RF 2017 Runtime]()

![OG 2018 Runtime]()
![RF 2018 Runtime]()

## Summary

### Advantages 
A general advantage of refactoring code is improved performance and increased scalability. In this example, we see a slight improvement to run time in the refactored code. For this instance, the run time only improved by a fraction of a second. 

If the data set was larger, the number of ticker symbols was greater, or both, the improvement would be more obvious. In the original code, the code loops through each row of the data set 12 times for each ticker symbol. In the refactored code, each row of the data set is only looped through once. If the number of rows in the data set where to increase, the difference in run time would become more significant. 

### Disadvantages  
A general disadvantage of refactoring code is the amount of time it takes to rewrite working code. This requires cost benefit analysis. Does the amount of time save of greater value than the time spent refactoring the code? For this instance, the benefit is minimal. If we know that the number of rows or number of ticker symbols will not increase drastically, it would likely not be worth a developer’s time to refactor the code.  
