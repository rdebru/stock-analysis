# **VBA of Wall Street**
	
## **Overview of Project** 
* Steve's parents are passionate about green energy. There are more green energy stocks to invest in, however they decided to invest all their money into DAQO New Energy Corp. Steve wants to help his parents and he decided to analyze a handful of green energy stocks in addition to DAQO's stocks as he is concerned about diversifying their funds. By automating and refactoring the analysis of the stock data, we will help Steve to reuse code with any stock and as well reduce chance of accidents and errors. We will also help Steve so that he will be able to see the results of the analysis with a click of button for all the stocks and for any year.
 
### Purpose 
* Edit and refactor Module2 VBA Script so we loop through the data one time and get all of the information for all the stocks. Also measure code performance between the refactored VBA script and the original script for the outputs of year 2017 and year 2018. As well to create a summary report based on the analysis and finding of the refactored VBA script compared to the original script of the module.

## **Results** 

![Elapsed run time for Output 2017](C:/Users/Ruth/OneDrive/Desktop/Class_Work/Module_2/HW2_Submission_ExcelVBA/Resources/VBA_Challenge_2017.png) 

![Elapsed run time for Output 2018](C:/Users/Ruth/OneDrive/Desktop/Class_Work/Module_2/HW2_Submission_ExcelVBA/Resources/VBA_Challenge_2018.png) 

* we can see from the images above that the time to run the refactored VBA Script for the outputs of year 2017 and year 2018 is faster than original script of the module. The run time for year 2017 is (0.12109 seconds) and for year 2018 is (0.125 seconds) using the refactored code. Using original scriptor the run time for year 2017 is (0.8593 seconds) and year 2018 is (0.8632 seconds). There is a difference of approximately 0.63 seconds. The images for the elapsed run time of the Original script for year 2017 and year 2018 are on `All Stocks Analysis` worksheet. As we can see when we click on the button on `All Stocks Analysis` worksheet, the stock analysis outputs for years 2017 and 2018 are same for both refactored VBA script and Original VBA script. Eventhough the original scrip functions the same(outputs are the same) but it is slower than the refactored code for both years 2017 and 2018. There are two buttons on `All Stocks Analysis`, `Run Analysis for All Stocks refactored` button for refactored script and `Run Analysis for All Stocks` button for the original script.
 
* to refactor the original VBA script, we will use four arrays, the tickers array and three output arrays (tickerVolumes, tickerStartingPrices and tickerEndingPrices). The tickerIndex variable is used to access the index across all four arrays. First for loop is used to initialize the tickerVolume. The tickerIndex is initialize to zero before the for loop. One for loop is used to loop over all the rows in the spread sheet. `For i = RowStart To RowCount`. Inside the loop, increase the volume for current ticker (the selected tickerIndex) for each row. Then an if statement is used to check if the current row is first row of ticker index and the previous row's ticker does not match with the selected tickerIndex. If true, get starting price for current ticker(Selected ticker index).
`If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then`

* Then Second if statement is used to check if the current row is the last row with the selected ticker and the next row's ticker doesnot match. If statement is true, then get the ending price and increment the tickerIndex. then this is the end of the first ticker and the outputs are stored in the four array with zero index. 
`If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then`

* Then we access the second ticker by the incremented tickerIndex and we loop through rows and output the data for the second ticker to the output arrays. and we repeat this for 10 times till the last index of the arrays.

* Then we use for loop `For i = 0 To 11`to go through the arrays (where we stored the results), to output the Ticker, Total Daily Volume, and Return on to `All Stocks Analysis` worksheet. Then formatting is applied to show how each stock performed by changing the cell color. There is one for loop and the code is shorter and easier to follow. These code can be used for as many stocks as we want by changing the size of the arrays. Steve should be able to analyze data for any year, he just has to add the data of the stocks for the year he wants to analyze on new worksheet.


## **Summary** 

1. What are the advantages or disadvantages of refactoring code? 

* Advantages  
	* the code is efficient—by taking fewer steps and using less 	memory 
	* help Saved time and money in the future
	* by improving the logic of the code, it is easier for 		future users to understand or read.
	* helps to find exting Bugs that make the code slow 
	* code is fresh and runs faster
	* Clean code is much easier to update and improve
	* less complex and easier to maintain and scale
	* important exercise to remove code smell
		
* Disadvantages
	* if refactoring is not done properly, it might introduce new bugs and might not be effient.It becomes more complicated instead of simpler and easy to read 
	* You can potentially introduce bugs that your tests or exsiting tests won't catch 
	* It could be time Consuming: You may have no idea how much time it may take to complete the process
	* resources to refactor if refactoring is not successful
	
2. How do these pros and cons apply to refactoring the original VBA script? 
* The refactored original VBA script is faster and more efficient than the original script as we can see from the time it takes to run(Images above). This will become very usefull when we have large data to analyze(more than 3013 rows). 
* The logic of the code is improved by using tickerIndex and three output arrays. One for loop is used to loop through the data, so are going through the data once by incrementing the tickerIndex to access the tickres and storing them in the output arrays.
* It is easier to maintain and easier for future users to read. If we have lots of stocks more than 12 stocks to analyze, we can easily do so by increasing the size of our arrays.
* the Refactored code is shorter and cleaner 

* The cons of refactoring do not apply to refactoring the original VBA script in this case, since we can see from the run time elapsed and testing that the code is efficient, faster and improved. So, refactoring the code makes sense. Time spent is reasonable compared to results and faster outputs of larger data of stocks. As well No new bugs are introduced. 
