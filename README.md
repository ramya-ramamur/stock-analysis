# Analysis of stocks for investment
##### An analysis using VBA macros.
# Overview of Project
Our client Steve' parents want to invest in Alternative energies companies, and have decided to put all their money in one stock : "DAQO New Energy Corporation", Ticker Symbol: 'DQ" that makes silicon wafers for solar panels. Their decision is based on emotion rather than facts and research. So, Steve wants find about the stocks performance and a handful of other green energy stocks as he wants his parents to diversfy their portfolio. 

## Purpose
Steve has created an excel sheet with green energy stock performance with opening and closing price, and stock volumes for the years of 2017 and 2018. We will be performing the analysis using Excel and Visual Basic Application (VBA) macros. Steve wants to be able analyze stock performance for any year (including ones he might add in the future than only for 2017 and 2018.

## Data Cleaning
Data was cleaned using excel to remove white spaces, formatting according to data type and sorting in the ascending order of the ticker symbol.

### Results
Initial analysis was done first for the stock with ticker symbol "DQ" that Steves parents are interested in and then for all stocks in one year - 2018. That code was then refactored for analysing stocks for any year and a timer to evaluate if the processing time had improved. 

##### Step 1
First a dialogue box opens that asks for the year to be analyzed 

Code : yearValue = InputBox("What year would you like to run the analysis on?" & vbCrLf & ("2017 or 2018?"))

![Screen Shot 2021-10-30 at 11 38 04 AM](https://user-images.githubusercontent.com/75961057/139554873-5d8a0269-8102-472a-8d90-310060f5e3db.png)

##### Step 2
All stocks ticker symbols are assigned to variables and stored in an array. 

![Screen Shot 2021-10-30 at 11 43 36 AM](https://user-images.githubusercontent.com/75961057/139555000-c87f0261-1708-4425-b307-ab43792bd374.png)

##### Step 3
Created an array for Stock Volume, Starting and Ending Prices

![Screen Shot 2021-10-30 at 11 46 01 AM](https://user-images.githubusercontent.com/75961057/139555081-dba3b7c2-b3b1-4d0f-979e-58caba82fb3f.png)

##### Step 4
Now using the tickerIndex, we parse the worksheets only once instead of iterating through for each ticker symbol like we did in the initial analysis.

![Screen Shot 2021-10-30 at 11 46 52 AM](https://user-images.githubusercontent.com/75961057/139555155-ed72ebbf-b22f-498d-be49-a4419dc3099e.png)

##### Step 5
Output the results on a new worksheet named "Analysis of all stocks". 

<img width="228" alt="Screen Shot 2021-10-29 at 4 45 27 PM" src="https://user-images.githubusercontent.com/75961057/139555216-2c6c4c2e-50e0-4804-bbf5-9781b4c07059.png">

<img width="225" alt="Screen Shot 2021-10-29 at 4 46 01 PM" src="https://user-images.githubusercontent.com/75961057/139555220-f95d547e-bbd5-4b7a-97d9-1d5938b3bd8d.png">

