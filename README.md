# Stock Analysis with VBA

## Overview of Project
This analysis was created using VBA to help a friend, Steve, determine whether certain stocks are worth investing in. Using data collected from the years 2017 and 2018, an initial script of code was created to pull information into the “All Stocks Analysis” worksheet. This first script analyzed all stocks from the data sets, finding the total volume of trades for each stock, as well as those stocks’ returns. Once the “All Stocks Analysis” sheet properly listed and formatted the results, a second script was started to refactor the existing code. The goal of this refactoring was to decrease the code’s run time and make it easier for others to follow when viewing it for the first time.  

## Results
### Comparing Performance by Year
After running the VBA module for both 2017 and 2018, it is apparent which year these stocks performed the best. Below are images of each year’s stock performances for reference:
#### 2017
[Insert Image here]
#### 2018
[Insert Image here]
2017 clearly has more green cells in the “Return” column, meaning that more of these stocks had a positive return over the course of that year. In fact, only one of the stocks had a negative return that year. In 2018, however, there was a drastic change, as only two of the stocks had a positive return that year. Although only two years of data were analyzed, it is evident that these stocks are all rather volatile. Perhaps an analysis of more than just two years of data would better inform Steve’s investment decisions.

### Execution times of Original Script vs Refactored Script
The original script finds the value of the `Variables` we are looking for by running through all rows of the data a total of 12 times using `Nested For Loops`. The code outputs the desired data and runs in about 0.71 seconds, but there was room for improvement via refactoring. See the original `Nested For Loop` below:
[Insert Image Here]

The challenge of refactoring the script was figuring out how to change it so that it runs faster than it already did. As was stated earlier, the script runs through all lines of data 12 times, once for each ticker. But what if there’s a way to output the analysis while only having to run through the data once? That is exactly what the refactored script does. Rather than using `Variables` to pull the output data, the refactored script incorporates `Arrays` to do the same thing. Using `Arrays` tells the script to pull output data for all tickers in just the first run-through of the data set. That way there’s no need to run through the data set 12 whole times. See below for the refactored script:
[Insert Image Here]

Notice that the output values from the original script, such as `totalVolume`, `startingPrice`, and `endingPrice` were changed to `tickerVolumes()`, `tickerStartingPrices()`, and `tickerEndingPrices()`. Those edits are where the `Variables` were changed to `Arrays`. 

Also notice how the `For Loops` are separated into three individual loops. This means that the script is no longer telling the second `For Loop` to run those 12 times because it is no longer being run inside of the first `For Loop`. 

This refactored script runs in about 0.12 seconds, almost 1/7th the time of the original. See below for the official run-times before and after refactoring:
2017
Original
[Insert Image Here]
Refactored
[Insert Image Here]
2018
Original
[Insert Image Here]
Refactored
[Insert Image Here]
## Summary

### Advantages and Disadvantages of Refactoring Code
#### Advantages
Often times, the first code that is written isn’t necessarily the most efficient, so refactoring can help improve that efficiency. For example, in this analysis, the refactored code helped decrease the run-time of the whole script. Although the differences in run-time were marginal in this example, there could be a much larger difference had there been more lines of data to analyze. So, this improved efficiency means that now the code is better prepared to analyze larger data sets without needing to take several seconds to do so.

Additionally, refactored code can make the script easier to follow for others who might be viewing the analysis for the first time. There are normally several ways to write code that will all reach the same end goal. So even though the first draft of the code might produce the intended results, there could be a shorter and more efficient way to write it. Not to mention, if there is less code in the script it is easier to update (and debug if necessary), just in case it needs to be applied to a similar analysis that has different arrays, variables, etc.
#### Disadvantages
One of the disadvantages of refactoring is that it is time consuming. In this analysis, refactoring the script took a good amount of time to complete, and all it really did was decrease the run-time by a few tenths of a second. If this code is intended only for this type of analysis and will only use data sets of this size, then it doesn’t make much sense to refactor it. Why spend all that time just to save fractions of a second? The tricky part is that it is hard to tell how long the refactoring process will actually take before it is started.

Another disadvantage of refactoring is that it can be risky if it is not done properly. If someone is refactoring code that they didn’t originally write, or perhaps wrote a few months or years ago, it is hard for them to fully know/remember why each line of code was written the way it was. Without understanding how the code runs and what the end goal is, someone might consolidate lines that need to remain separate, or delete lines that seem redundant but actually serve an important function. 

### Advantages and Disadvantages of Refactoring this Stock Analysis
As was seen earlier in this written analysis, the original code uses `Nested For Loops` and `Variables` while the refactored code uses three separate `For Loops` with `Arrays`. Both scripts perform the same function, and they each have their own advantages and disadvantages. 
#### Advantages
Again, as was seen earlier, the refactored code runs the analysis quicker than the original code. So, in this case, refactoring resulted in a more efficient script that is better equipped for processing larger data sets.
Also, the refactored code is easier to update and a bit easier to follow. Since there aren’t `nested for loops`, any one `for loop` can be updated without affecting the others. This also clarifies the script, allowing other people to follow each statement without having to worry if it is embedded in a different, larger statement.

#### Disadvantages
As was detailed earlier, refactoring the code for this specific analysis only saved a few tenths of a second. Throughout the refactoring process, there was a lot of trial and error to make sure the code ran the same as it originally did. Even with knowing how the original code was created, it took several test runs to debug the refactored code so that it ran properly. So, really the one disadvantage of refactoring in this case was that it took a while to complete while only reducing run time by a rather insignificant amount.

