# Stock Analysis with VBA

## Overview of Project  
<br />


### Background:
In Module 2, we performed stock analysis on 12 stocks to determine total volumes and returns with VBA. This analysis logic was iterative where we looped thru the dataset 12 times, once for each stock, to calculate returns. The next step is to expand the research to entire stock market, and while the iterative approach worked well for 12 stocks, it may not perform well when the analysis is expanded to all stocks. 
<br />
<br /> 

### Scope:
The scope of this project is to:
 - Performance tune the code so that it runs faster by refactoring in a way that it loops thru the dataset once and collects the same information. 
- Analyze 1) how run times compare between iterative and refactored code and 2) advantages and disadvantages of code refactor.  
<br /> 
<br /> 



## Analysis and Results
The focus of this challenge is to improve and analyze the performance of the VBA code. VBA module on the spreadsheet (*VBA_Challenge.xlsm*) has the following subroutines of interest:

-	AllStockAnalysis() --> this contains iterative code which loops thru entire dataset for each ticker.
 - AllStocksAnalysisRefactored() --> refactored code which loops thru dataset one time.  

 The analysis provides the following features:  

- Separate buttons to run and clear analysis
- Ability to input the year for analysis
- Color formatting to distinguish position and negative returns
- Reports total esplased time taken by the script. 

<br /> 

 ### Conclusions:  
 
 1) Looking at 2017 elapsed time stats, we can see that there is over 99% improvement with the refectored code that eliminates looping logic per ticker. Please refer to elapsed time screenshots representing improved run times using refactored code

 ![](/resources/VBA_Challenge_2017.png)  

 ![](/resources/VBA_Challenge_2018.png)


 2) Another observation is that the first run of the subroutines seem to take a little longer, possibly due to initial resource allocation in memory. But subsequent runs show more or less the same elapsed time for 2018 and 2017.  

 3) Additional potential improvment can be to declare and populate *tickers* array dynamically (instead of hardcoding ticker values) which will become very useful when the analysis is expanded to all stocks. Doing so may cause elaspsed time to go up slightly (from 0.26 sec to 0.32 for 2017 which is a 20% increase using subrountine *AllStocksAnalysisRefactoredWithDynTickerArray()* ) but the code scalability that it offers outweighs the slight increase in run time.  

![](/Dynamic_Population_of_Tickers_Array.png)

<br /> 


## Summary
Overall, the purpose of refactoring is to organize/structure the code in a better way but not change its behavior. Listed below are some of the pros and cons of this approach. 

### Advantages of Code Refactor:
- It can help improve the design of application which can:
    - Impove the overall performamce of the code, such as with stock analysis application where removing the looping logic resulted in considerable improvement in runtime of the script.
    - Make future maintenance easier because of improved design
- The process can identify bugs missed in the intial development of the application

### Disadvantages of Code Refactor:
- While refactor can help identify missed bugs, it can also introduce bugs if: 
    - Proper test cases are not established. 
    - The application is large and code structure is complicated  

- The process requires proper resource/time allocation and thorough testing of larger applications.

For example, in case of stock analysis, addition of code to dynamically populate tickers array required peroper testing to make sure all values were captured. 

There are pros and cons of code refactor and each application should be assessed for areas of improvement, time/resource availability, project dealines and cost associated before making any major changes.