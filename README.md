# Stock-Analysis

### Project Overview

The objective of this project was to analyze stock data using VBA macros in microsoft excel. The stock data contained data for two different years 2017 and 2018 for 12 different energy related stock tickers looking to find volume traded over the year and percentage return. The goal was to give the end user the ability to see certain analysis done so by using macro's to analyze the data and embedding macro buttons for the end user to input a select year and clear the sheet. After doing so the code was reformatted to make the macro's analyze the data quicker and allowing the macro to be quicker on a larger scale.

### Results 

**Stock Analysis Results**

With the running of the macro's for the years 2017 and 2018 the following stock analysis results were yielded.

![image](https://user-images.githubusercontent.com/85713568/136495839-e79694eb-78b0-4e8e-a92f-354374df99cc.png)
![image](https://user-images.githubusercontent.com/85713568/136495906-f548a9e1-cf06-4c4e-a2ea-5ae83a7544ea.png)

From the images above we are able to tell that in 2017 most stocks had a positive return and were in the green with the exception of TERP being down 7.2%.
For 2018 we see that most stocks were down except for ENPH and RUN with those two being positive atleast 8x% and positive the year before.

**Code Refactor Time Difference**

The following pictures show the time difference for the macro's to execute for each year.

2017 original code                                                        
![image](https://user-images.githubusercontent.com/85713568/136496649-6190e494-2d39-4df2-bd83-d93a3f2ca4af.png)

2017 refactored code

![image](https://user-images.githubusercontent.com/85713568/136497045-8a89f9ef-ae14-44f5-a8f9-d37d96850d0c.png)


2018 original code

![image](https://user-images.githubusercontent.com/85713568/136497151-13cef933-502a-4b0a-a241-4e21b1f47a0a.png)

2018 refactored code  
![image](https://user-images.githubusercontent.com/85713568/136497191-887164cb-0bfa-4c20-9e14-aef8de0a4e1a.png)

The original code for 2017 and 2018 both took over half a second where as the refactored code took .07 seconds to run. The refactored code took under 20% of the time to complete the analysis versus the original code. 
The difference between the original and the refactored code is that the refactored code loops through all the data only once filling the data for each ticker into an array which updates the ticker as the ticker data changes as the rows go on and prints the data after the data has been sorted. While the original code sorts through the entire list once for each ticker and prints the data after each ticker.


![image](https://user-images.githubusercontent.com/85713568/136498176-d07a840f-6090-4004-86d7-aa3d28aaccfc.png)
Here is the original code indicating that all the rows were ran through each time for every new ticker.

![image](https://user-images.githubusercontent.com/85713568/136498310-70ceea5a-96df-41c1-a134-41993769b98c.png)
This refactored code shows how a ticker index array was used to move along the tickers as the data was combed through once rather than repeating itself for every ticker.

### Summary

In summary refactoring code may allow code to become more efficient after laying down the foundation and flow with the original code. Although refactoring may become time consuming and sometimes cause inefficiencies.
For example in this refactored vba script it ran at under 20% of the original codes time. If the data were much larger containing thousands of different stock data it would take much longer for the macro to run and loop through the data everytime with the original code, whereas the refactored code increasing the ticker index at every new ticker would make it more seamless. Although refactoring the code did take some additional to figure out a more efficient approach to sorting the data and whether or not the refactoring would have been more efficient is not guaranteed. 

Both codes were dependent on the stock data having each individual ticker data for the year bunched together and sorted by date. The flaw in the refactored code is if it were applied to data which was not already sorted. If the stocks were not sorted properly by the ticker index variable would not be able to properly increase the volume and find the appropriate starting and ending dates as it would increase the ticker index moving onto the next ticker. In the case of improperly sorted ticker data the original code would still be able to find the proper ticker since it would search through the entire lis for each ticker and adjust the data. In both cases both codes would not work unless the code was atleast sorted by date.
