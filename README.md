# Stock Analysis through VBA

Coding in VBA to analyze the volume and return of select green energy stocks. 

## Overview of Project

Client's family believes in investing in green energy stocks due to the finite supply of fossil fuel as basis for energy. Currently, his family has all of their investments in DAQO New Energy Corp, a manufacturer of silicon wafers for solar panels. Client would like to diversify his investment portfolio in green energy stocks to increase the chances of return on their invesments. Client would like to assess the success of each company's stock in order to inform his decision in selecting candidates for diversifying his investment portfolio. A unique **ticker** was assigned as an abbreviated identifier for each company. 

### Purpose

A stock analysis of green energy companies from the years 2017 and 2018 was performed to visualize the total volume as well as the return (based on starting and ending prices) of each company. This analysis was performed using VBA Macros to automate the analysis process to encompass individual stocks for both years. The initial VBA Macro code was written focusing on the analysis of ticker *DQ* for DAQO New Energy Corp. The design pattern of this code was used as a basis to build the code to include all tickers. Finally, this code was refactored to improve the run-time for all stocks analysis based on year.

## Results

The analyzed results were printed in table format based on the ticker and its associated total volume and return. The return was derived from calculating the percent increase or decrease of the stock based on its starting and ending prices. Conditional formatting was applied to the return column to differentiate between increase and decrease. For each analysis done, the runtime of the code was displayed in conjunction to evaluate the performance of the code. 

### Analysis of Stock Performance 

![Stock Analysis for Year 2017 And Runtime for Refactored AllStocksAnalysis Code](https://github.com/Fabalin/stock_analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

Based on the analysis performed for the year 2017, *DQ* boasts the highest return rate of all the stocks at 199%. Almost all stocks showed an increase in return, save for *TERP* which experienced a 7.2% decrease in return. In contrast however, *DQ* has the lowest Daily volume total of around 36 million. Compared to the highest daily volume of 782 million of *SPWR*, *DQ's* stock experiences less trading per day and therefore *DQ* is the most volitile stock compared to the rest. This is because the low daily trading volume total indicates that on average, a trade of this stock would have more of an impact on its returns. The runtime for the refactored code was within a hundreth of a second.  

![Stock Analysis for Year 2018 And Runtime for Refactored AllStocksAnalysis Code](https://github.com/Fabalin/stock_analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

Based on the analysis performed for the year 2018, *DQ* boasts the lowest return rate of all the stocks at -62.6%. In contrast to 2017, almost all stocks in 2018 showed a decrease in return, save for *ENPH* and *RUN* which experienced an increase of around 80% respectively. *ENPH* and *RUN* also had the first and third highest daily volume totals meaning that their stocks were the least volitile out of the rest. The code for this analysis ran slightly slower but still within the hundredth of a second.

#### Which Other Stocks to Invest

Based on the analysis *DQ* is a volitile stock to invest all of the client's capital. Hence, it is important to diversify the portfolio by spreading this investment across other stocks. Judging from the Daily volumes and returns for both years, *ENPH* and *RUN* are great candidates since their returns remained as increase from 2017 to 2018 and their volumes remained robust for both years. 2018 appeared to be a bad year for green energy stocks since we saw a decrease in return for most companies. Despite this, *ENPH* and *RUN* still maintained their postiive returns. *RUN* in particular saw an increase in their returns from 2017. 

### Analysis of Code 

Globally, the script would cycle through the rows of the dataset defined by its year and would print the analyzed values in a different workbook based on the parameters specified. The data would then be formatted to help visualize the differences between each stock. For loops were used to cycle through the dataset and conditionals were used to extract the data for analysis if the criteria for tickers is met. 

#### Original Code

The original code for the all stocks analysis ran by looping through the total rows of data from the top of the data set to the bottom, based on the ticker provided. In order to find the bottom of the dataset, sample code was utilized. Sample code to find the last row of the data set taken from [StackOverflow](https://stackoverflow.com/questions/27784658/use-the-last-row-count-in-a-formula)
```
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
```
Since there were 12 tickers in total, an array of 12 tickers stored as a string data were created. Another loop was constructed and overlayed on top of the initial loop to cycle through the rows. This created a nested for loop that would allow the program to scroll through all the rows multiple times based on the number index of the tickers array.  

![Runtime of Original AllStocksAnalysis Code](https://github.com/Fabalin/stock_analysis/blob/main/Resources/GreenstocksOGTime.PNG)

This code ran close to half of a second and is a five fold increase in runtime from the refactored code. This is due to the number of loops utilized to run the analysis and cycle through all the rows of data multiple times. 

#### Refactored Code

The refactored code features a single loop that cycles through the rows of data once. to complete the analysis and uses the ticker index variable as a reference that unifies the 4 arrays created. These arrays are the tickers, total volume, starting and ending prices. Each array stores the data specific to the ticker that is identified by the ticker index variable. Finally, after storing the data based on the ticker idenfied, the ticker index variable would be increased to shift the analysis and store the analyzed data into a different index of the total volume, starting and ending prices arrays. This would effectively allow the program to loop through the data once to complete the analysis for all tickers. 

## Summary

### Advantages or Disadvantages of Refactoring Code

Refactoring code allows the user to redefine how the issue is resolved through further refinement. This increases the efficiency of the code performance and allows the system to maintain more memory to run other programs efficiently. Furthermore, refactoring allows the user to gain better comprehension of ones' own code and improve upon it to aid others in understanding the code written through simplification. However, refactoring is a process of revision and edition. This process could introduce new bugs and errors into the code and thus could potentially create more problems to solve. 

### Relating to Refactoring The Orginal VBA Script

Despite the longer runtime, the original script was thorough with its analysis. The script ran through the rows multiple times for each ticker so this could work in cases where the dataset is not sorted and organized based on the tickers. The refactored script capitalizes on the fact that the datasets provided were sorted and grouped based on tickers, to cycle through the data once and collect the information specified. The process of refactoring the original script involved multiple instaces of debugging as to double check that the new arrays created were storing the data specificed. The ```Debug.Print``` command was crucial in monitoring the status of the arrays as it progresses through the code. Additionally, since the refactoring utilized the old code, its structure can limit the creation of the new refactored code if the idea for the refactoring isn't properly clarified. Outlining the key differences by comparing and contrasting the design patterns of the old and new code is important to prevent getting stuck and lost in the old code. 

