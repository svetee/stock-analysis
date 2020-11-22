# STOCK ANALYSIS WITH VBA 

## OVERVIEW: VBA Stock Analysis Project

### Purpose of Project

The purpose of this analysis was to edit, or refactor, the provided Stock Market Dataset with VBA solution code to loop through the entire the data one time to collect an entire dataset. In result we refactored the code to make the VBA script run faster. The goal was to make the code run faster by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results: Refactor VBA Code & Measure Performance

### Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:

1. The tickerIndex is set equal to zero before looping over the rows.

Ticker index variable was created and set equal to zero before applying to all the rown. The tickerindex is used to access the correct index across the four different arrays on VBA Code (tickers array and  three output arrays created). 

![The tickerIndex](https://user-images.githubusercontent.com/60243906/99915803-ebacd600-2ca9-11eb-803f-bc0ed94a1f45.png)

2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

![Arrays are created](https://user-images.githubusercontent.com/60243906/99915810-fbc4b580-2ca9-11eb-900d-6987182d769c.png)

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

![The tickerIndex is used to access the stock ticker](https://user-images.githubusercontent.com/60243906/99915831-1f87fb80-2caa-11eb-8694-44576402b837.png)

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

![The script loops through stock data, reading and storing](https://user-images.githubusercontent.com/60243906/99915839-30387180-2caa-11eb-9416-bb6beea889d2.png)

Stored values from tickerStartingPrices and tickerEndingPrices

![Start and EndPrices Code](https://user-images.githubusercontent.com/60243906/99915843-3fb7ba80-2caa-11eb-8a44-ccbcb7f355fb.png)

5. Code for formatting the cells in the spreadsheet is working.

By highlighting positive returns green and negative returns red, it is easier to determine which stocks did well and which ones didn't. 

![Code for formatting the cells in the spreadsheet is working](https://user-images.githubusercontent.com/60243906/99915848-4ba37c80-2caa-11eb-88ed-1ad00394f450.png)

6. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

![Dataset examples provided](https://user-images.githubusercontent.com/60243906/99915859-60801000-2caa-11eb-9eec-9938cf528202.png)

![Dataset examples provided (1)](https://user-images.githubusercontent.com/60243906/99915863-6e359580-2caa-11eb-88d6-0953dcef914d.png)


Below the Final VBA Analysis screenshots,

![VBA_Challenge_2017](https://user-images.githubusercontent.com/60243906/99915873-7ee60b80-2caa-11eb-9858-53ba7bf1abc3.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/60243906/99915874-81486580-2caa-11eb-96f6-873850c2caa0.png)

7. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

![Time for 2017 analysis](https://user-images.githubusercontent.com/60243906/99915885-958c6280-2caa-11eb-89cf-bc1ae84860e9.png)
![Time for 2018 analysis](https://user-images.githubusercontent.com/60243906/99915886-96bd8f80-2caa-11eb-9d2c-ff0a37c52ada.png)

## Summary

### Deliverable with detail analysis:

#### Advantages or disadvantages of refactoring code in general

Refactoring code needs to happen in small steps. It's achieved through tiny changes in program, each step improving the performance. 

##### Disadvantages:

1. A long code might be repetative, through logic, duplicate lines can be eliminated.  
2. A logical structure may be duplicated in two or more procedures 
3. It is better to split a complex unstructured code into several functions.
4. Refactoring can affect the testing outcomes if not done carefully

##### Advantages:

1. Logical errors easily appear in well structure code built with nested conditionals and loops.
2. Using Excel flow displays program logic easily
3. VBA interpretation (Excel) of code can show patterns easily overlooked inthe source.

#### Advantages and disadvantages of the original and refactored VBA script

"Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. Now, let's think about something, What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand? If yes then definitely you didn’t pay attention to improve your code or to restructure your code."

It is important to consider the code refactoring process as cleaning up our house. Too much clutter in a home is chaotic and stressful. - The same goes for written code. Finally, a well-organized code is easy to change, to understand, and to maintain. By paying attention to the code refactoring process earlier one can avoid avoid facing difficulty later. 










