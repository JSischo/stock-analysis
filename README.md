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

### Comparsion of Execution Times of Original Macro and Refactored Macro

[Speed of 2017 analysis with refactored macro.](/Resources/VBA_Challenge_2017.png)

> The link above is a screen shot of the run time for one attempt of the 2017 stock data utilizing the refactored code.

[Speed of 2018 analysis with refactord macro.](/Resources/VBA_Challenge_2018.png)

> The link above is a screen shot of the run time for one attempt of the 2017 stock data utilizing the refactored code.

## Summary

- What are two conclusions you can draw about the Outcomes based on Launch Date?
  - There seems to be an increased chance of a successful Theater campaign when it is launched in late spring or early summer. Specifically May, June and July. On the flip side, December does seem to be the worst time to start a campign if you would like it to be successful. My hypothesis here is that in December, people tend to be a little tighter with their cash as they are buying gifts for others during the holiday season. 
  - There does not seem to be a specific time of year correlation between failed campaigns and when they were launched. The trend line seems to be somewhat flat through the entirety of the year. There may be other factors to look at to help determine what variables are correlated with failed campaigns. 

- What can you conclude about the Outcomes based on Goals?
  - The data analysis seems to indicate a bimodal result when looking at successful campaigns based on their funding goals. Campaigns that were seeking < $1000 dollars were the most successful at nearly 80% followed closely be campaigns asking between $35,000 and $44,999.

- What are some limitations of this dataset?
  - One main limitation is that the n count for campaigns that had a higher goal amount is very low. Only 42 campaigns had a goal of >$25,000 (4% of total campaigns), so it is hard to formulate strong conclusions for the higher goal campaigns.
  - Another limitation is that the "Theater" filter includes some very high value goals, many of which were unsuccessful, because they were campaigns looking to build physical structures and NOT simply to be able to support an isolated production. We should probably filter these out to strengthen our analysis.

- What are some other possible tables and/or graphs that we could create?

  - Other tables/graphs that might give us further insight into what may be correlated with S/F campaigns:
  
   1. Percentage funded. We could created new "Successful" benchmark based on how close a campaign was to reaching it's goal.
   2. Number of backers per campaign. Was success related to # of donors?
   3. Avg. donations. Did the successful campaigns have a higher avg./backer?
   4. Re-run current analysis with a Country filter specific to the country that Louise plans to launch her play.
