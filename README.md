# Stock Analysis with VBA

## Overview of Project

Refactor existing code to analyze data for a number of stocks:
- Refactor previous macro to enhance proficiency
- Add Arrays to the macro
- Compare the time from the refactored macro to the previous macro

### Purpose

  The purpose of this project was to to create and refactor database analysis macro in VBA for our client, Steve, to help him analyze a large stock database from multiple years of data as quickly as possible. This project included a **refactoring** of our original macro which included many ***For*** loops and ***If...Then*** statements. This required using and array approach to the project. We also included a **User Input** box and **Run Analysis** and **Clear Data** buttons to assist Steve in simplifying the user interface with his database file.
  
  Steve will be able to use our final product to assist his clients by providing historical stock performance information to help guide their future stock purchase decisions.

## Results

### Comparison of 2017 vs. 2018 Stock Performance

When we run an annual stock comparison between 2017 and 2018 for the 12 stocks in this database, a number of things stand out:
* Overall, 2017 seemed to be a very good year for many stocks as 11 of the 12 stocks posted positive returns in their annual price per share. One quarter of the stocks achieved gains between 100% to 200% in their price. The only stock posting a loss was TERP.
* On the other hand, 2018 wasn't as favorable to the stock prices of our 12 stock sample. 10 of 12 posted negative gains in their price per share. Not as significant as the positive gains of 2017, but some stocks lost around 60% of their value.
* There were 2 stocks that posted year over year gains, ENPH and RUN.
* Overall total volumes varied year to year with no apparent correlation between volumes and gains/losses.

### Comparsion of Execution Times of Original Macro and Refactored Macro

[Speed of 2017 analysis with refactored macro.](/Resources/VBA_Challenge_2017.png)

> The link above is a screen shot of the run time for one attempt of the 2017 stock data utilizing the refactored code.

[Speed of 2018 analysis with refactord macro.](/Resources/VBA_Challenge_2018.png)

> The link above is a screen shot of the run time for one attempt of the 2017 stock data utilizing the refactored code.

## Summary

By completing this assignment, it became apparent that there some advantages to refactoring code, which include:
1. A functional macro already exists to perform the task that you are looking to perform, thus saving some time in creating a more efficient one.
2. If is somewhat easy to see where the existing code can be adjusted by looking for the ***For*** loops and nested ***If...Then*** statements.
3. The refactored code usually required less commands in general by decreasing repetition within the original code. 
4. The new code is typically more efficient and performs it's function much quicker than the original code.

There are a couple disadvantages of refactoring code that I encoutered during this project:
1. It took some time to determine how to refactor some of the original code and adjust the commands to utilize arrays.
2. New bugs can be introduced which will require some time to debug and correct.

For this specific project, there was a definite advantage in macro execution time between the original code and the refactored code. This was discussed previously in the Results section above. Another advantage that I observed was that the refactored code was much cleaner and easier to follow as some of the loops and If...Then statments could be removed, thus making the final code easier to follow. 

Some of the disadvantages that I encoutered was that it did take a fair amount of time to figure out what could be changed to improve the script performance. I attribute most of this to my unfamiliarity with VBA and coding as this is very new to me. During the refactoring process I also found that the basic functionality of the original code changed due to inconsistencies or errors in my refactored code, thus creating new bugs. This took some time to debug and get the new code to do what I needed it to do. So if time is a constraint, refactoring a properly functioning code may not be the best option.
