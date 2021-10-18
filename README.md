# Module-2-Stock-Analysis

# Project Overview

The purpose of this excercise was essentially to streamline the performance of our code to produce the same result using less computer power and in a shorter period of time. Through a process called refactoring, we were able to make the code more efficient in order to better suit Steve's needs of potenitally running this analysis on a much wider range of stocks.

# Results

A basic assessment of the results of the analysis run for the years 2017 and 2018 on these 12 stocks clearly reveals that the market fared much better in 2017 than it did the following year. In 2017, all but one (TERP) of the 12 stocks analyzed produced a positive return, whereas in 2018, all but two (ENPH, RUN) produced negative returns. The use of conditional formatting in our final charts makes this discrepancy immediately obvious, as seen below.

![2017Result](https://user-images.githubusercontent.com/91569387/137675893-f8c243f5-5beb-4f73-b566-d0ca5cec2fe0.png)
![2018Result](https://user-images.githubusercontent.com/91569387/137675908-eb9af017-5c56-4e86-8105-20dcb295e8b6.png)

For an even more nuanced analysis of these findings, charts could be created that analyse and display the overall performance of each stock over both years weighted in proportion to the total volume of the stock over the same period. This would provide a clearer picture of which of these 12 stocks would likely be the most reliable investment for Steve and his parents.

In terms of the performance of the code, the refactoring seems to have significantly improved the speed and efficiency of running the analysis. In refactoring, we've improved the code so that it only sweeps through the large data set once and in doing so notes all of the data will will need. Previously, the code swept through the entire set three separate times: once to retrieve the total volume for every iteration, again to determine the starting prices, and once more to determine the ending prices of each stock. Thus, this streamlined process takes less time and memory to run than the earlier verison of the code because it works through each iteration in the large data set once rather than three separate times. This improved efficiency is evident via VBA's timer function. The following are the respective times to run the 2017 and 2018 analyses with the original code:

![2017time](https://user-images.githubusercontent.com/91569387/137678776-03399a53-d5da-4fcf-89e1-330aa869da71.png)
![2018time](https://user-images.githubusercontent.com/91569387/137678787-f84f1213-a889-4d79-884e-470d1038be55.PNG)

The following are the respective times to run the 2017 and 2018 analyses with the refactored code:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91569387/137678835-41285653-1353-4617-826e-cb0d264401b1.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/91569387/137678849-e9fb0b45-962e-4b4c-a756-b57a0d8b3f69.PNG)

Clearly, the refactored code is much faster, reducing the times to run the analysis by about 68.5% for 2017, and about 80.8% for 2018.
