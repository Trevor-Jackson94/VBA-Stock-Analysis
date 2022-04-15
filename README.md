#Stock Analysis Refactoring Report

##Overview
For this analysis we created code and refactored that code for Steve to be able to quickly analyze the *Total Daily Volume* and *Return* for a variety of stock data and their tickers. We accomplished this by creating a macro to run in Microsoft Excel using Visual Basic for Applications coding language. Steve needed a larger group of stocks to analyze to give his parents a better investment than their original stock of choice. This data set can be updated easily and can be viewed on a yearly basis as well.

## Results 

### Run Time Improvements
After changing the code to use arrays for the *tickerVolumes()*, *tickerStartingPrices()*, and *tickerEndingPrices()* along side the *tickerIndex* variable, we successfully refactored the *AllStocksAnalysis* code that we started with. This new code improves the run time of the macro significantly. Previously using the *AllStocksAnalysis* code we were able to run the analysis without any spreadsheet formatting in 0.78125 seconds (Figure 1). 

![Figure 1: Macro Run Time Using *AllStocksAnalysis* Code for 2018 Data](https://github.com/Trevor-Jackson94/VBA-Stock-Analysis/blob/main/Resources/Previous%20Time%202018.PNG)

After refactoring the code we were able to improve the run time to 0.1328125 seconds (Figure 2).

![Figure 2: Macro Run Time Using *AllStocksAnalysisRefactored* Code for 2018 Data](https://github.com/Trevor-Jackson94/VBA-Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

The same can be seen with the 2017 data but even more significantly. The 2017 data improved from 0.7695313 seconds to 0.09375 seconds using the new refactored code (Figure 3 and 4).

![Figure 3: Macro Run Time Using *AllStocksAnalysis* Code for 2017 Data](https://github.com/Trevor-Jackson94/VBA-Stock-Analysis/blob/main/Resources/Previous%20Time%202017.PNG)

![Figure 4: Macro Run Time Using *AllStocksAnalysisRefactored* Code for 2017 Data](https://github.com/Trevor-Jackson94/VBA-Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

### Refactored Coding Changes
The decrease in run time was accomplished by using arrays in the For loops instead of multiple variables. In the refactored code we used three (3) new arrays including a tickerVolumes() array, tickerStartingPrices() array, and tickerEndingPrices(). We also used a tickerIndex variable that increases with each For loop to use to dictate which entry in the arrays we are referring to (Figure 5).

![Figure 5: Refactored Code showing tickerIndex variable and New Arrays](https://github.com/Trevor-Jackson94/VBA-Stock-Analysis/blob/main/Resources/Refactored%20Code.PNG)

## Summary
To conclude, the *AllStocksAnalysis* VBA code was refactored to generate the new *AllStocksAnalysisRefactored* VBA code that runs as a macro in Microsoft Excel to analyze stock data that includes multiple stocks and their tickers. We added a tickerIndex variable to use within the For loops as an index in the three (3) new arrays we created tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices(). Adding these arrays and the tickerIndex variable decreased the run time for the 2017 and 2018 data significantly dropping them from 0.7695313 seconds to 0.09375 seconds and 0.78125 seconds to 0.1328125 respectively. The advantages to refactoring code is helping to optimize and simiplify the code. Some disadvantages are there are chances you can introduce new bugs into the code and the code may not be as versatile after the refactoring. In our situation we accomplished all the pros to refactoring our original VBA code. Along the way, we did run into a variety of bugs that had to be fixed, but in the end we have a very versatile code that can work efficiently for much larger sets of data. A draw back to our VBA code is when adding more data and new tickers we will have to edit the ticker array with the names and new number of tickers. 
