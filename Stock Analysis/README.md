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

### Results

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

## Code Analysis
	Code Analysis section will describe the sections of the code written to achieve specific goal. It will have multiple 
	sections detailing what each piece of code is doing.	
	
### Code for checking the year provided and returning the error message.
	To return the error, I used the loop to check the input year against the name of all the sheets in 
	the Workbook and if there's no matching name found then return the error messge exit out of the code.
	Check the in-line comments for the logic behind teh code
### Image - VBS Code for Error Checking
![VBA_Year Check Erro Message](https://user-images.githubusercontent.com/98173091/153690612-c12af768-3ebe-4080-a0fa-bac317189669.png)

### Code for 




	Year 2017 was quite good for investors who invested in any of the stocks from the list except for "TERP". 
	The average return for 2017 was 67.3%, with SEDG and DQ best performing stocks.

	
	After the intial analysis of the data it appeared to be just a stored data not providing much of intelligent 
	or actionable information. 
	We knew we could get the information we wanted using 
		- pivot tables and charts
		- fucntions like years, sum, vlookup and countif etc.
	Once the data was formatted and new data created using fucntions, pivot table and charts were inserted into new sheets 
	we could draw conclusions.
	
	CHALLENGES
	1. "Deadline" and "Launched" columns had just numbers and i did not have any idea about what data it is.
	It was only after using the date functions we could make that data readable.
	2. COUNTIFS, had to google the syntax to check for data in multiple columns. Like in "Outcomes based on Goals" excercise for 
	all Goals >= 1000 and <4999 and subcategory ="Play" and outcomes ="Successfull" or "failed".

### Image - Challenge COUNTIF Complexity 
![COUNTIFS](https://user-images.githubusercontent.com/98173091/152464471-42e87666-47d0-4e92-9fc7-10779d466cee.png)

### Image - Understanding Dates in Pivot Table Row Areas
![PIVOTAREAS](https://user-images.githubusercontent.com/98173091/152464601-34699487-0565-4667-9a07-73d7ee0221c6.png)


### Code for checking the year provided and returning the error message.
	To return the error, I used the loop to check the input year against the name of all the sheets in 
	the Workbook and if there's no matching name found then return the error messge exit out of the code.
	Check the in-line comments for the logic behind teh code
### Image - VBS Code for Error Checking


	The pivot table for outcomes based on launch dates was easy to draw from the parent sheet "Kickstarter". Only thing we need to do 
	was select and add the appropriate columns into the folowing pivot table areas
		1. Filters - Added Parent Cateroy and Years as filters
		2. Rows - Added Date Created
		3. Columns - outcomes because we wanted to see succesfull, failed and canceled as columns
		4. Values - added outcomes, becuase we wanted to see the count of each outcome.
	
	CHALLENGES
	Possible challenges could be understaing the additon of Year, Quarter in Row section while adding the "Date Created" into 
	Row Section	

### Image - Outcomes Based on Launch Date
![Theater_Outcomes_vs_Launch](https://user-images.githubusercontent.com/98173091/152464393-e8007763-2644-431e-8613-f9580d6b35d5.png)


### Analysis of Outcomes Based on Goals
	
	This was exciting becuase all the data we need for Outcomes Based on Goals was derived using functions COUNTIFS, SUM etc. 
	It was straight forward column selection and inserting a chart. After creating the colums and labels for Goal, 
	# of Successful, Failed and Canceled Outcomes and % of f Successful, Failed and Canceled Outcomes. It was all about 
	using the appropriate function.
	
	CHALLENGES
	1. COUNTIFS, had to google the syntax to check for data in multiple columns. Like in "Outcomes based on Goals" excercise for 
	all Goals >= 1000 and <4999 and subcategory ="Play" and outcomes ="Successfull" or "failed".
	
### Image - Outcome based on Goals
![Outcomes_vs_Goals](https://user-images.githubusercontent.com/98173091/152457665-ac4499f8-2567-4e9e-84f2-f252ad82c443.png)
	
### Challenges and Difficulties Encountered

	CHALLENGES
	Challenges and under the analysis section.
	
	DIFFICULTIES
	1. Biggest hurdle was to work with GitHub, i'm still struggling.
	2. Managing work and studies. 

### Results

#### What are two conclusions you can draw about the Outcomes based on Launch Date?

	1. Campaigns launched in May,Jun and Jul in all years are most successfull for not just "Theater" but for 
	all categories as well. This is telling us people tend to go out and spend time to watch theatre during summers which is 
	also the time for school summer vacations.
	2. Campaigns launched (Theater or all Categories) (after mid year) towards the year end tend to fail, this could 
	be because of ending of yearly budget and year end vacations.

#### What can you conclude about the Outcomes based on Goals?

	For Category "Plays" - Smaller the Goals, Higher is the Success Rate. 
	As we increase campaigns' funding goal to highr values it tend to fail more, with only exception of 
	funding goal of $30,000 - $45,000 where it saw a success rate of 66%. Goals of $45000 and above are 
	absolute failure(only 12% success), we can divert our efforts towards campaigns wih goals of $0- $20000 range which has 
	success rate of more than 50% 
	
#### What are some limitations of this dataset?
	
	The dataset was limited to just 4114 data points (sample size was small) when we consider the # of countries involved.

#### What are some other possible tables and/or graphs that we could create?
	
	We could have created more tables and graphs to determine 
		1. Country specific Outcomes for each category
		2. Something like Top 5 countries where "Theater" is preferred over other categories
		3. Campaigns where pledged funding fell short by just less than 20% of its goal.
		


