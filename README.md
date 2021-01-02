# stock-analysis
stock analysis for Steve
# Kickstarting with Excel

## Overview of Project

An Analysis of Kickstarter funding campaigns with the following focus/criteria:
- Filtered for "Theater" campaigns
- Outcomes based on "Launch date" and "Funding Goal"

### Purpose

  The purpose of this endeavor was to assist our client, Louise, to determine if there seemed to be any correlation between successful(S)/failed(F) Kickstarter campaigns in the **Theater** category. We looked at two outcomes per her request: Campaign ***Launch date*** and ***funding goal***.
  
  Louise will use this information to help guide her in planning a funding campaign for her play *Fever*.

## Analysis and Challenges

### Analysis of Outcomes Based on Launch Date

[Outcomes vs. Launch Date](/Resources/Theater_Outcomes_vs_Launch.png)

> The link above contains a graph which analyzes the success and failure of Theater based kickstarter campaigns and when during the year the campaign is launched. The results do seem to indicate that there is an advantage to achieving a successful campaign based on when during the year the campaign is launched. 

### Analysis of Outcomes Based on Goals

[Outcomes vs. Goal](/Resources/Outcomes_vs_Goals.png)

> The link above contians a graph that displays the relationship between successful/failed campaigns relative to the amount of the fund raising goal of the campaign. There does seem to be some corrlation between these two variables.

### Challenges and Difficulties Encountered

The only difficulty that I encountered during this challenge was that pivot tables are a new Excel concept to me, so it did take some practice and playing around with the functionality to achieve a level of confidence with creating the appropriate tables and associated charts. 

## Results

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
