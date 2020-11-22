# STOCK ANALYSIS WITH VBA 

## OVERVIEW: VBA Stock Analysis Project

### Purpose of Project

The purpose of this analysis was to edit, or refactor, the provided Stock Market Dataset with VBA solution code to loop through the entire the data one time to collect an entire dataset. In result we refactored the code to make the VBA script run faster. The goal was to make the code run faster by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results: Refactor VBA Code & Measure Performance

### Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:

1. The tickerIndex is set equal to zero before looping over the rows.

Ticker index variable was created and set equal to zero before applying to all the rown. The tickerindex is used to access the correct index across the four different arrays on VBA Code (tickers array and  three output arrays created). 

Picture the ticker index 

2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Picture arrays were created

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

Tickerindex picture here

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

picture the script looks throuch stock data

Stored values from tickerStartingPrices and tickerEndingPrices

picture start and end price codes

5. Code for formatting the cells in the spreadsheet is working.

By highlighting positive returns green and negative returns red, it is easier to determine which stocks did well and which ones didn't. 

Code for formatting cells

6. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

example pictures


Below the Final VBA Analysis screenshots,

7. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png












