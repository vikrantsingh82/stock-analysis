# Stock Analysis

## Overview of Project

	Steve has got two years of stock data for numerous tickers. He want to anayze the yearly stock performance
	for multiple tickers (companies) to determine best and 	worst performing stocks for years 2017 and 2018 at 
	the click of a button without losing too much time, not using exel funtions which is time sonsuming.  .
	
	Steve should be able to run the stock analysis not only for years 2017 or 2018 but more. If more data is added 
	to the worksheet with more tickers for another year, only the button click wil be enough to run the analysis 
	and present the data in tabular format for each year.

### Purpose

	The pupose of the analysis is to draw conclusions based on yearly perfomance of stocks to identify
		a. stocks to invest
		b. stocks to sell

## Analysis and Challenges
	
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


### Analysis of Outcomes Based on Launch Date

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
		


