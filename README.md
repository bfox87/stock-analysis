# Stock Analysis Using Excel VBA

## Project Overview

### Background:
An initial analysis of a dozen stock returns by year was performed using Excel VBA code to determine which stocks appeared best to invest in. The original code proved satisfactory for the initial goal. However, to expand the analysis to thousands of different stocks, the initial code may be slower and take up more computing power than necessary. More efficient code can be used to create a more efficient analysis that is beneficial when working with a much larger dataset.

### Purpose:
To refactor the original code in order to improve the efficiency or calculation time needed to generate the returns for thousands of stocks.

## Results

### Analysis of Outcomes Based on Launch Date:
To analyze campaign outcomes by launch date, a pivot table and pivot line chart were used. The data was filtered by theater and year and organized to show the number of successful, failed, and canceled campaigns by month. The chart below shows all years of theater data with each line a different campaign outcome. The chart was saved as picture and linked to the Resources folder in this repo to be included in the analysis.



### Analysis of Outcomes Based on Goals:
A line chart was used to analyze campaign outcomes by fundraising goal $ amounts. Some initial summary data was compiled first through the use of countifs functions. To begin, goal $ amount categories like (<1,000, 1,000-4,999) were created. Then countifs formulas enabled me to get a numerical breakdown of campaign outcomes by goal buckets and play subcategory. Finally, percentages of the total were used to compare the outcomes. This serves as a better metric than total number of successful/failed outcomes. The chart below shows this percentage breakdown between outcome categories as they track over fundraising goal $ amounts. The chart was also saved as picture and linked to the Resources folder in this repo to be included in the analysis.



### Challenges and Difficulties Encountered:

Testing VBA code:
```
'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
```

I encountered some difficulty putting the countifs formulas together as I was getting an error from Excel saying “There’s a problem with this formula.” I had forgotten to specify the $ goal criteria in brackets as I figured numbers didn’t need quotes like text did. A bit of “google-fu” solved the problem! I also had issues getting the Outcomes vs Goals image link to work on this readme file. Appears a typo was the cause.

## Summary

### Advantages and Disadvantages of Refactoring Code
To 

### Advantages and Disadvantages of Refactoring Stock Analysis Code

As it pertains to this project, 

![VBA_Challenge_2017](https://github.com/bfox87/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

too

![VBA_Challenge_2018](https://github.com/bfox87/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

T00
- Outcomes vs Launch Date Conclusions:
    1. When all years are looked at together, the best time for a theater kickstarter campaign is late Spring/early Summer. This is the time of year when the likelihood of success is highest. The month of May appeared the most popular month for total theater kickstarters with the number of successful campaigns particularly pronounced. Roughly 2/3 of the total campaigns launched in May were successful.
    2. It appears to be a bad decision to launch a theater kickstarter in the month of December. The number of successes and failures are roughly even. This makes logical sense as most people are busy with holiday festivities and can be financially strapped based on other gift or donation commitments common during the holidays.

- Outcomes vs Goals Conclusions:
    1. It is recommended to keep your fundraising goals modest to give your campaign the highest likelihood of success. Close to 3 out of 4 fundraisers with goals under $5,000 are successful.

- Dataset Limitations:
    - Using this dataset in 2022 to analyze outcomes by launch date raises some concerns. There appears to be no theater data beyond the year of 2017. This is five years ago and things may have changed since then. Ideally more recent data would be included.
    - Additionally, the location data could be more detailed than country level. Some of these countries are quite large (i.e. U.S.) and there might be significant variability in fundraiser characteristics and outcomes due to geographic region, particularly city the play will be held in.
