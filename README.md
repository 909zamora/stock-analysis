# stock-analysis

## Purpose
The purpose of this project was to analyze stock performance in 2017 & 2018 without having to do any manual inputs besides the VBA codes. In this assignment I was analyzing stocks in order to have a recommended stock for my parents. The analysis took into account volume trading as well as 11 different stocks but was modeled so it could handle thousands of stocks at a time. At the end of the day, the model illustrated in red or green, the performance of the stock for the year and also returned a run time to see how long it took to run. 


## Results
From the analysis, it was easy to conclude that the stock "DQ" was not a good pick for the parents. The stock DQ had a -62% return for 2018 and I would recommend investing in the stock "ENPH" as they had a great year in 2017 at 130% return & an 82% return in 2018. This was the stock with the highest return and DQ had a -199% return in 2017 so its historical performance was troubled from the beginning. DQ also had the lowest trading volume and lower volumes means that there are fewer shares being traded during the year. This would also hint that "DQ" is less liquid than the other stocks on the list. Low volume could also mean increased risk with big sell offs from one major investor causing the price to be very volatile.

## Advantages & Disadvantages of Refactoring Code
The advantages for refactoring the original code is that it made the VBA more efficient, the run time was now closer to 0.11 seconds as opposed to 0.65 seconds in the beginning. This decreased time might seem very minimal but it makes a huge difference when you are running thousands of stocks as opposed to just the 11 we did in this analysis. It would make the VBA run 6x faster, meaning if it took 6 minutes to process a large data on the original VBA, the refactored one would only take 1 minute. Time is a huge advantage but also the coding gets improved as the firs time a code is made, it is usually not the best. It needs to be refined, just as you refine baking your favorite cake over time. You learn the timing, the placement of certain variables/ingredients makes a huge difference.

The disadvantages of refactoring the code is that, you could cause an error that breaks the original code or it can run and still work but have a missed step along the way, simply put, it is subject to more human. It also takes more time so sometimes you need to prioritize what are the needs of the business/project because if the code is working well and you don't anticipate the code to break with future data sets, then why change it to begin with?

## Advantages & Disadvantages of the Original and refactored VBA script
The advantages of the original vba script was that it was very explicit on the steps for just those 11 stocks. So if you only wanted to do an analysis on the stocks that you selected, option one would probably be better because it can hand pick what you want. The analysis was done specifically to analyze DQ and the other 11 stocks but is not ready for more which is a disadvantage.

The advantages of the refactored VBA script is that it will be ready to execute a bigger analysis on more stocks, it will always run faster than the original VBA script since its code is more efficient at running. The disadvantage of the refactored code is that it now how more variables, we added "tickerVolumes" / "tickerStartingPrice" / "tickerEndingPrice" / etc. so it wil be harder for someone coming in to get comfortable or understand fully. The other code is easier to read since it is simpler and has less assigned variables. 



End Sub
