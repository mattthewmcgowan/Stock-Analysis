# Stock-Analysis

## Project Overview 

The primary purpose of the analysis was to utilize VBA in the analysis of stock. The analysis then expanded from a single year to multiple years of data. From there it expanded even further to include mutiple stocks thus requiring new macros to accomodate for numerous ticker symbols. Ultimately, the purpose of this VBA project was to give our client a working sheet for all their future stock analysis at the click of a button! 

## Results 

When I was initially handed this project, my client only wanted me to analyze the stock performance of one stock over one year. 

``` Range("A1").Value = "DAQO (Ticker: DQ)" ```

```  If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ"  ```

After the completion of this first task I was able to loop through each row, looking their specified ticker symbol (DQ) and find the return percentage. 

Then the question was asked, "Well, what stock is viable?" 

I was able to rewrite the code to include multiple years by adding in a range loop that will loop through all rows of both sheets of data: ```Range("A1").Value = "All Stocks (" + yearValue + ")"```

Then I built a message box that will prompt the user to determine which years data they would like to run the macro for. 

Finally, once the loop has run through the new array of ticker symbols the results are formatted to let the client know green or red to reflect the ticker symbols who've returned a positive or negative yearly percentage.  

Although the client model was working I went back through and refactored my code structre to make the macro run smoother. 

### Reslutes from the final refactored VBA analysis

<img width="365" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/106042900/173966954-c2c402db-ad0c-4b15-b4cd-8cfb2f11a1fc.png">


<img width="365" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/106042900/173967006-1dd4f72c-a7e2-4e20-b7d2-f7f5b173fb8f.png">

## Summary 
What are the advantages and disadvantages of refactoring code? 
  - Advantages: The advantages of refactoring code would be finding ways to make any macro run smoother. It is always good practice to make your code as simple as possible and try not to repeat yourself. Refactoring code will put you in a mindset to make your statements as ledgible as possible. Refactoring also gives you the oppurtunity to make your code neat and readable for whoever might come behind you. 
  - Disadvantages: when refactoring your code you will most likely run into errors. The errors aren't in your logic but in the way you are presenting your logic to VBA. It can be a frustrating process to refactor a script that already works. 

Both the disadvantages and advantages apply when refactoring the orginal VBA script. The original script can often be clunky and poorly formatted. Taking the time to refactor the script will make it more legible and run faster. 
