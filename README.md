# stock-analysis

## Purpose
The purpose of this project was to use VBA code to analyze the performance of the stock market over 2017 or 2018. We  refactored the code we had previously used. This means we had to tweak the code a bit so that we were able to use it more efficiently. In the code itself used the index function to easily go through a lot of data that otherwise would have been very time consuming.  By doing this we could analyze the annual percentage change in stock.

## Results
Overall, the stock market performed better in 2017, with returns as high as 199%.  In 2018 the market performed poorly. VBA gave me a lot of trouble with step 4:

Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) – 1
 
While I cannot find the issue if I remove this, I am able to get the timer to work. However, my numbers are incorrect. Images can be found here:

     stock-analysis/Challenge 2 Resources/Challenge 2 Resources/VBA Challenge 2018.png

     stock-analysis/Challenge 2 Resources/Challenge 2 Resources/VBA Challenge 2017.png

## Summary
While refactoring may seem to make things easier, for me there were a lot of issues in debugging my codes. The code worked for me in the original form.  I spent hours trying to figure out the bugs, and was able to solve all but one (above). I could not find any solution. In refactoring you still have to plan out your code and think about what each steps means. It was a cleaner and easier way of coding compared to the original code we used. 
