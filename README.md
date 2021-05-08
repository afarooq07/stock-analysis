# Stock Analysis with VBA

## Overview of Project  
<br />


### Background:
In Module 2, we performed stock analysis on 12 stocks to determine total volumes and returns with VBA. This analysis logic was iterative where we looped thru the dataset 12 times, once for each stock, to calculate returns. The next step is to expand the research to entire stock market, and while the iterative approach worked well for 12 stocks, it may not perform well when the analysis is expanded to all stocks. 
<br />
<br /> 

### Scope:
The scope of this project is to:
1) Performance tune the code by refactoring so that it loops thru the dataset once and collects the same information. 
2) Compare run times between iterative and refactored code
3) Evaluate advantages and disadvantages of code refactor.  
<br /> 
<br /> 



## Analysis and Results
The focus of this challenge is to improve and analyze the performance of stock analysis VBA code. The analysis consists of the following deliverables:  

1) Separate buttons to run and clear analysis
2) Ability to input the year to calculate total daily volumns and retun for each stock
3) Color formatting to distinguish positive and negative returns
4) Comparison of total esplased time taken by the script. 

<br /> 

 ### Resources: 
 1) Datasource: VBA_Challenge.xlsm
 2) Images: Available under resources folder:
    - VBA_Challenge_2018
    - VBA_Challenge_2017
 3) Code Details - Subroutines of interest:
    - AllStockAnalysis() --> this contains iterative code which loops thru entire dataset for each ticker.
    - AllStocksAnalysisRefactored() --> refactored code which loops thru dataset one time.
 
 <br />

 ### Conclusions:  
 
 1) Looking at 2017 elapsed time stats, we can see that there is over 99% improvement with the refectored code that eliminates looping logic per ticker. Please refer to screenshots below (left: after refactor, right: before refector):  

 <br /> 
 <img src="/resources/VBA_Challenge_2017.png" width=300 align=left> 
 <img src="/resources/2017_ElaspsedTime_Without_Refactor.png" width=300 align=center>
 <br />

 2) The first run of the subroutines seem to take a little longer, possibly due to initial resource allocation in memory. But subsequent runs show more or less the same elapsed time for 2018 and 2017.  

 3) Additional potential improvment can be to declare and populate *tickers* array dynamically (instead of hardcoding ticker values). This will make the script generic and applicable on any number of stocks in the same file format. 
     - The sample code below explains the concept (*subroutine: AllStocksAnalysisRefactoredWithDynTickerArray*)
     - Please note that there may be a better way to accomplish this via a built-in function but I couldnt find a reliable example online. 


<img src="/resources/Dynamic_Population_of_Tickers_Array.png" width=700 align=center>

<br /> 


## Summary
Overall, the purpose of refactoring is to organize/structure the code in a better way but not change its behavior. Listed below are some of the pros and cons of this approach. 

### Advantages of Code Refactor:
- Code refactor is vital step in SDLC, it can make development process efficient by first coding the logic and then organizing it the best way possible rather than spending a lot of time on design upfront which can hinder the progress and delay deliverables. 
- It can help improve the design of application which can:
    - Impove the overall performamce of the code. In stock analysis, removing the looping logic resulted in considerable improvement in runtime of the script.
    - Make future maintenance easier because of improved design.
- The process can identify bugs missed in the intial development of the application.

### Disadvantages of Code Refactor:
- While refactor can help identify missed bugs, it can also introduce bugs if: 
    - Proper test cases are not established. 
    - The application is large and code structure is complicated  
- The process requires proper resource/time allocation and thorough testing of larger applications. For example, in case of stock analysis, code refactor and addition of logic to dynamically populate tickers array required proper testing to make sure all values were captured. 

There are pros and cons of code refactor and each application should be assessed for areas of improvement, time/resource availability, project dealines and cost associated before making any major changes.
