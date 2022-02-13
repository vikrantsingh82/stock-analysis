# Stock Analysis with VBA

## Overview of Project
	Steve has got two years of stock data for numerous tickers. He want to anayze the yearly stock performance
	for multiple tickers (companies) to determine best and worst performing stocks for years 2017 and 2018 at 
	the click of the button withpout losing any time, not using exel funtions which is time sonsuming.  
	
	Steve should be able to run the stock analysis not only for years 2017 or 2018 but more. If more data is added 
	to the worksheet with more tickers for another year, only the button click wil be enough to run the analysis 
	and present the data in tabular format for each year. The VBS code should also display the time taken to proces the data
	for each year.

### Purpose
	The pupose of the project is to refactor the VBA code we have been using to learn loops, conditions, arrays etc. in VBA 
	to make it more efficient—by taking fewer steps, using less memory, and improving the logic of the code. 		

## Results

### Using the worksheet to run the analysis.
	User would need to open the VBA_Challenge.xlms sheet and go to "AllStockAnalysis" sheet, provide a year and run the
	stock anaylsis. User is required to click on "Run Stock Analysis" button and the provide the year.
	
#### Image - All Stock Anaylsis Sheet
![VBA_Challenge_Initial Scree](https://user-images.githubusercontent.com/98173091/153690032-9a7ec62a-43b6-49f2-8df2-1535f9ac98fa.png)
	
	Click on the "Run Stock Analysis" button and provide year 2015 or 2016. (NOT 2017 or 2018), it should return a error message
#### Image - Wrong Input
![VBA_Challenge_Input_Wrong_year](https://user-images.githubusercontent.com/98173091/153690318-9b276073-360f-4e51-b48b-795220eb058b.png)

#### Image - Error Mesasge if year provided is not in the workbook
![VBA_Challenge_Error_Message](https://user-images.githubusercontent.com/98173091/153690342-c3d790bf-c332-4c6c-84ba-dd8c7531656e.png)

### Stock Analysis for the Year 2017
	
	Year 2017 was quite good for investors who invested in any of the stocks from the list except for "TERP". 
	The average return for 2017 was 67.3%, with SEDG and DQ best performing stocks.
	To run the stock analysis, Click on Run Stock Analysis button and provide the year 2017 in the input box.
	
### Image - Providing year 2017 as Input 	
![VBA_Challenge_Input_Pop_up](https://user-images.githubusercontent.com/98173091/153690863-5ce974be-3c5d-46be-aa1a-eab15c772926.png)

### Image - Output after the stock anaylis code executed for year 2017
![VBA-Challenge_2017_Analysis](https://user-images.githubusercontent.com/98173091/153690899-1f5b8cc5-6070-40e6-b85f-eb3f582951b5.png)

### Image - Clearing 2017 Analysis Data
	Before running the stock analysis for year 2018, we would like to clear the data for year 2017 .
	To clear the data user need to click on "Clear" button,
![VBA-CHallenge_Clear_Cells](https://user-images.githubusercontent.com/98173091/153691029-35ca13ac-c786-4863-b337-be79b686f0c7.png)


### Stock Analysis for the Year 2018
	After running the analysis for Year 2017 and clering the data, run the stock analysis for the 2018. 
	The average return for 2018 was -8.5%, with ENPH and RUN best performing stocks.
	To run the stock analysis, Click on Run Stock Analysis button and provide the year 2018 in the input box.
	
### Image - Providing year 2018 as Input 	
![VBA_Challenge_Input_2018](https://user-images.githubusercontent.com/98173091/153727397-f61dc1aa-a53f-4d58-8d2f-5944922d07d8.png)


### Image - Output after the stock anaylis code executed for year 2018
![VBA-Challenge_2018_Analysis](https://user-images.githubusercontent.com/98173091/153727415-30de9d36-54b3-4015-8fbc-0b823c2ed2b0.png)

## Code Analysis - Start
	Code Analysis section will describe the sections of the code written to achieve specific goal. It will have multiple 
	sections detailing what each piece of code is doing.	
	
### Code for checking the year provided and returning the error message.
	To return the error, I used the loop to check the input year against the name of all the sheets in 
	the Workbook and if there's no matching name found then return the error messge exit out of the code.
	Check the in-line comments for the logic behind teh code
### Image - VBS Code for Error Checking
![VBA_Year Check Erro Message](https://user-images.githubusercontent.com/98173091/153690612-c12af768-3ebe-4080-a0fa-bac317189669.png)

### Code for the timer - Start and End and adding column headers in "AllStockAnalysis" sheet.
	Timer code is used to calculate the total execution time.
![VBA_Challenge_Start_End_Timer](https://user-images.githubusercontent.com/98173091/153729140-d161ad3e-e68e-41df-8908-ebb7de7ed1d4.png)
![VBA_Challenge_Adding_Col_Headers](https://user-images.githubusercontent.com/98173091/153732259-25d7fbf8-ae4f-4841-a4a2-95e5c36274b9.png)


### Code to declare a dynamic arrays and using count of unique tickres to size them.
	Instead of declaring ticker array with size 11, i decided to calculate the count of unique tickers and 
	assign it to a variable/ SO that we can use this variable to initialize the array.
	For this i created a new VBA fucntion 'CountOfUniqueTickers' which was taking the year provided as arguments, used the for
	loop to iterate thriugh each record in Ticker column and add unique values to a list ibject and then reurn the list.count 	
![VBA_Challenge_Count_Of_Unique_Tickers](https://user-images.githubusercontent.com/98173091/153729385-56ab48c0-d8cb-4b74-aab5-240fe12ab8f0.png)

	I declared the dynamic string array variable, and then had to use ReDim statement to size  or resize a dynamic array. 
	In this case i we will used the variable that holds the count of unique tickers.
![VBA_Challenge_Initialize_Ticker_array](https://user-images.githubusercontent.com/98173091/153732381-e3b5804a-814c-4b09-aea9-9256f948a87a.png)

	Similar to dynamic ticker array we need to declare other dymanic arrays for volume, start and ending price.
![VBA_Challenge_Initialize_Other_arrays](https://user-images.githubusercontent.com/98173091/153732511-f1ca3610-dd63-4178-991d-2cffcb343385.png)

### Code to populate the ticker array with unique tickers.
	Instead of hard coding the array values, i used the for loop to populate unique ticker values
	' NOT Planning to use this hard coded array
	    ' tickers(0) = "AY"
	    ' tickers(1) = "CSIQ"
	    ' tickers(2) = "DQ"
	    ' tickers(3) = "ENPH"
	    ' tickers(4) = "FSLR"
	    ' tickers(5) = "HASI"
	    ' tickers(6) = "JKS"
	    ' tickers(7) = "RUN"
	    ' tickers(8) = "SEDG"
	    ' tickers(9) = "SPWR"
	    ' tickers(10) = "TERP"
	    ' tickers(11) = "VSLR"
	See the code below, also instead using the column number for ticker, colume and close columns in year sheet,
	i decided to programatically get the column index.
![VBA_Challenge_Populating_Ticker_Array](https://user-images.githubusercontent.com/98173091/153732644-b7627b83-3afc-468b-adbc-052801f17d3d.png)

### Code to calculate Total Volume, Start and End Price
	Used the for loop to iterate through each ticker in year sheet, caculated the total volume, 
	starting and ending price for each ticker. And then assigned the calculated values to thr three
	dynamic arrays we declared earlier.
	
	tickerVolumes
        tickerStartingPrices
        tickerEndingPrices	
	Please note that variable tickerIndex is set to zero before we stared the for loop.	
![VBA_Challenge_Assign_Vol_Return_each_Ticker](https://user-images.githubusercontent.com/98173091/153733318-01a6fea1-1ffe-4044-9e58-8493f20a1865.png)
	
### Code to read and print the calculated data from dynamic arrays
	Once the total volume, start and end price is calculated and stored in its respective dynamic array, we need to 
	read each each in a loop and assign the values to Cells in "AllStockAnalysis" sheet to display end result.
![VBA_Challenge_Read_And_Populate_Data_Using_loop](https://user-images.githubusercontent.com/98173091/153733392-1e9a9978-9d3f-4e82-93df-9537ddbe7a96.png)

### Code to format the output
	After all teh calculations are done and data is read from teh arrays, we need to format the data, 
	for example, if teh return is greater than 0 then color the cell in Green else Red. And make the 
	header names appear in Bold.
	In addition to this, i merged the three columns and colored it in yello for the top row "All Stocks (year)"
![VBA_Challenge_Formatting](https://user-images.githubusercontent.com/98173091/153733589-f83996f2-4a93-4596-a21a-f827033e0878.png)

## Code Analysis - Completed

## Summary
	After refactoring the code we used for learning loops, arrays and conditions, we can clearly see the performance 
	improved considerably. See the images below to compare the performnace of the original and refactored code.
### Execution Time : Refactored Code
	After refactoring the code and using appropriate loops and condtion we noted the execution time for
	stock analysis for years 2017 and 2018 took close to a second.
![VBA-Challemge_2017_ExecutionTime - I](https://user-images.githubusercontent.com/98173091/153733813-cfb2813e-547a-481d-90bd-6500fd93a115.png)
![VBA-Challemge_2018_ExecutionTime - I](https://user-images.githubusercontent.com/98173091/153733822-1af8e77f-c68c-41e2-aa72-532248e90784.png)
	
	When we ran the refactored code for second time it took even less time becuase of resources and memory already in use.
![VBA-Challemge_2017_ExecutionTime - II](https://user-images.githubusercontent.com/98173091/153733836-dd90c776-ad92-4a35-b9b9-92a37e3d2b5f.png)
![VBA-Challemge_2018_ExecutionTime - II](https://user-images.githubusercontent.com/98173091/153733843-54cd438d-e388-471b-a2c5-156307360cde.png)
	
	
### Execution Time : Old Code 
	Old code when run for years 2017 and 2018 used to take more than 6 seconds to run the analysis.
![VBA-Challenge_2017_NOT Refactored Code](https://user-images.githubusercontent.com/98173091/153733877-4f955704-b82f-4527-964a-dc396fe4b835.png)
![VBA-Challenge_2018_NOT Refactored Code](https://user-images.githubusercontent.com/98173091/153733888-79694cdd-1b8a-428e-8b79-2b2230e2e016.png)

