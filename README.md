# VBA-of-Wall-Street
Use VBA  to analyze stock data

## Overview
The purpose of this project was to use VBA to analyze stock market data.  We had to create a for loop that would allow us to easaly change the year and show how long the analysis took.  Then we had to refactor our code to give us a faster running time.

## Results
As we can see in the tables below, for the 12 stocks we looked at, 2017 was a much better year.  The only stock that fell was TERP.  With a drecrease of only 7.2%, a reletivly small amount, if investors had spread their money evenly accross these stocks then they would still have had and increadible year, four of the stocks had a return of over 100%.

2018 was not a good year for these stocks.  Only two saw positive returns.  DQ, witch had the higest return in 2017, had the biggest drop in 2018.  RUN witch had the smalles positive return in 2017 did the best in 2018, with a 84% return.

![refactored 2017](https://user-images.githubusercontent.com/117960721/235834291-3cd221d6-56ba-4502-9a7f-5b3def7c6583.png)
![refactored 2018](https://user-images.githubusercontent.com/117960721/235834305-e5cbb06a-d5be-4127-a120-6c027f086d02.png)

In order to increase the output speed of the code I refactored it by creating a ticker index variable and three new output arrays.

![code 1](https://user-images.githubusercontent.com/117960721/235836482-44f2c0a3-18a3-4fc8-b868-e8b774e0d476.png)
![code 2](https://user-images.githubusercontent.com/117960721/235836491-b9ff0410-dec3-40ad-843d-68676f62751e.png)

The run time for both years decreased by about .6 seconds each.  Although this may not seen like alot, a decrease of this amount in such a small sample should translate to a decrease of quite a bit in a larger sample.

![time 1 2017](https://user-images.githubusercontent.com/117960721/235834351-f468ff31-460f-474d-a612-adc14ac3417d.png)
![time 2 2017](https://user-images.githubusercontent.com/117960721/235834368-5aee502b-f116-4c83-904a-6de4a1bb1132.png)


![time 1 2018](https://user-images.githubusercontent.com/117960721/235834391-c9a651a7-21f7-4032-86ab-e80d04e44456.png)
![time 2 2018](https://user-images.githubusercontent.com/117960721/235834413-9912a55a-48cd-4dde-9ab0-05808f78e388.png)


## Summary
I was able to show that refactoring the code does help to decrease the analysis run time.  With such a small sample size it did not really make sense to refactor the code.  It took me far longer to refactor the code then the difference in run times.  But with a larger sample size to analyse, and more years to look at, spending the time to refactor the code could potentialy save a huge amount of time.  The refactored VBA script, used in this project, does have the advantage of a more streamlined layout.  This would make it easier to look back on when creating another VBA script.  
