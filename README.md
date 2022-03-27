# stock-analysis2

MODULE 2: VBA_Challenge.vbs

![image](https://user-images.githubusercontent.com/90135381/158843476-bc128967-9a07-44e5-accc-1f55261eca86.png)

IMAGE: obtained from: https://www.pngwing.com/en/search?q=stock+market

_______________
***RESOURCES:*** excel/VBA.


_______________
***OVERVIEW:*** This challenge was generated to assist Steve with a stock market analysis for his parents using the VBS Stock Market Dataset, converting it to XLSM to, ultimately, expand the datset to the entire stock market over a period of years. To augment this analysis, Steve refactored the code to loop the same information in order to acquire the same result, hopefully, in a more efficient script run through, so that this can be applied to more stocks. Refactoring in this case is defined as follows: 

-not changing the data, but making the analysis more efficient (run through the script faster/fewer steps)

-utilizing less memory

-while improving the logic of the code


_____________
***RESULTS:***

***DELIVERABLE 1: Refactoe VBA Code and Measure Performance*** 

In this deliverable, the VBA_Script was refactored through date one time and the data was collected and displayed via:

**-1a:** a tickerIndex variable was created and set to equal zero prior to looping through the rows

**-1b:** 3 output arrays  (tickerVolumes/tickerStartingPrice/tickerEndingPrice) were created and separated into long and single data types

**-2a:** a for loop was created to initialize the tickerVolumes to zero

**-2b:** a for loop was created to loop over all rows/spreadsheet

**-3a:** current tickerVolumes were increased and the tickerVolume was added to the current stock ticker (tickerIndex=index)

**-3b:** the first row with selected tickerIndex was checked with an if-then statement and current starting price was added to tickerStartingPrices

**-3c:** another if-then statement was created to check closing price to tickerEndingPrices

**-3d:** a script was written to increase the tickerIndex if the following row's ticker doesn't match previous row

**-4:** a for loop was created to loop through the arrays from 1b to output the Ticker,” “Total Daily Volume,” and “Return” columns

**-FINALLY** buttons were created for ease of vizualization of the results for 2017 and 2018, with an associated pop-up window and a clear worksheet option
as illustrated in FIGURES: 1-4


<img width="1066" alt="Clear Worksheet" src="https://user-images.githubusercontent.com/90135381/160297640-d1dad9b2-0173-4ce9-94b8-ad3dcfbd8898.png">


                                            FIGURE 1: Clear Worksheet


<img width="1377" alt="Pop-up" src="https://user-images.githubusercontent.com/90135381/160297642-7baaccae-6553-47a0-984d-2c910e6af2fe.png">


                                            FIGURE 2: Pop-up window


<img width="1240" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/90135381/160297644-2ee512f2-1ba5-413c-bd79-1f846a8395a9.png">


                                            FIGURE 3: VBA Challenge 2017 PNG


<img width="1265" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/90135381/160297646-4ce973b0-70bb-4cc5-8fbb-0165bb2ac494.png">


                                            FIGURE 4: VBA Challenge 2018 PNG


***DELIVERQBLE 2: Written Analysis***

- repo with README

***SUMMARY:***
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
________________
***REFERENCES:*** Google, BSC, StackOverflow, GitHub
