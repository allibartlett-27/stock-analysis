# stock-analysis
Module 2 - VBA 

## Stock Analysis Results

#### **Overview of Project:**
###### *Purpose*	
The purpose of this Stock Analysis was to help Steve show his parents the performance of DAQO stock in the year 2017 and 2018, as well as other Green House Stock options his parents should consider. I also wanted Steve to be able to present his findings in a timely and presentable manner that was easy to read and understand. 
 
#### **Results:**
###### *2017 v. 2018*
Since Steve's parents were most interested in the Stock DAQO (DQ), let's start with that. DQ performed very well in 2017 with return of almost 200%, however is 2018 DQ 	had a negative return of 62.6%. This is true for almost all of the 12 stocks we analyzed. Only stocks ENPH and RUN both has positive returns in 2017 ans well as 2018. 
[VBA_Challenge_2017](https://github.com/allibartlett-27/stock-analysis/blob/main/VBA_Challenge_2017%20(Table).PNG) vs. [VBA_Challenge_2018](https://github.com/allibartlett-27/stock-analysis/blob/main/VBA_Challenge_2018%20(Table).PNG)

My recommendation would be to invest in ENPH as it was the highest daily traded stock and had a positive return of 81.9% in 2018. Both criteria that Steve's parents were looking for. 

Refactoring our original script also decreased the run time for both the 2017 and 2018 All Stock Analysis. By stream-lining and cleaning up our code our run time dropped from .8242188 seconds in 2017 and .9179688 seconds in 2018, to .1484375 seconds and .1445313 seconds. 
[2017 Original Script Runtime](https://github.com/allibartlett-27/stock-analysis/blob/main/green_stocks_2017.PNG) & [2017 Refactored Script Runtime](https://github.com/allibartlett-27/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
[2018 Original Script Runtime](https://github.com/allibartlett-27/stock-analysis/blob/main/green_stocks_2018.PNG) & [2018 Refactored Script Runtime](https://github.com/allibartlett-27/stock-analysis/blob/main/VBA_Challenge_2018.PNG)
  
#### **Summary:**
###### *What are the advantages or disadvantages of refactoring code?*
Refactoring gives you the opportunity to make a code more efficent and clean. Having another set of eyes look at it can help navigate errors and identify areas of improvement with a code. It seems as though refactoring can be a time-consuming process just to increase the process time by fractions of a second. I can see the benefit with a code that is looping thousands or millions of data sets through, as the efficiency would be greater. Possibly, another disadvantage could be refactoring a portion of your code could result is errors if other statements within that same code rely on a portion you are refactoring.  I am not sure exactly if this is a scenario that would happen as the code we refectored was realtively small but could image this may be an issue with more complicated codes that run multiple different parts within the same macro or subroutine. 

###### *How do these pros and cons apply to refactoring the original VBA script?*
As I was refactoring this code, I had to use the DeBug a couple of times to identify areas of the code that were affected by changes and updates I was making. It was more work to put in for a minimal upgrade in return of the information. However, when the refactored script was complete, it did run faster and looked good with the formatting also included nicely. 
