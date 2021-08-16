# Analysis of Steve's Stocks Using VBA

## Overview of Project
The purpose of this project was to expand our original data set to include the entire stock market over the last few years. I had to refactor the original Module 2 code for this challenge in order to determine which stocks Steve should invest in. My refactored code produces the Total Daily Volume and Return of many different stocks. 

## Results 

After running my refactored stock analysis, I was able to confirm that my results were the same as the original code. Below you can view the stock performance between 2017 and 2018:

<img width="305" alt="Screen Shot 2021-08-15 at 10 00 23 PM" src="https://user-images.githubusercontent.com/88108455/129501776-cba0b42f-a777-4eae-b7b8-673402d4d747.png">

For 2017, we can see that all of the stocks have a return above 0.00% except for one, TERP.

<img width="308" alt="Screen Shot 2021-08-15 at 9 59 41 PM" src="https://user-images.githubusercontent.com/88108455/129501848-2c78a624-273d-47c1-be9c-27d0d16478fd.png">

For 2018, we can see that only 2 stocks have a return above 0.00%, which indicates that 2017 was a much better year to invest. 

Below you can see the elapsed time for running the script on the years 2017 and 2018:

<img width="252" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/88108455/129625145-583c2aad-5b9c-45e5-9a58-5f99d0eb1f75.png">

The script took 0.148 seconds for the year 2017.

<img width="254" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/88108455/129625230-1bacf671-7134-4f7e-9a49-85c746efaf3a.png">

The script took 0.140 seconds for the year 2018.

Overall, after refactoring the code, the execution times prove to be much faster and more efficient than from results in the original script. The original script took over one second to complete while the refactored code takes under 0.2 seconds.

## Summary

1. The biggest advantage of refactoring this code was that the elapsed time decreased by almost 8x. A disadvantage is that refactoring the code actually takes a lot of time, however the end result is beneficial due to the fact that it runs faster and improves the user experience moving forward.
2. The pros apply to refactoring the original VBA script because our code is now cleaner and has a better design than the original green_stocks file. If a user were to compare both files side by side, our new refactored code is much easier to understand with the comments and code. The cons apply to refactoring the code as well because not everyone may have time to go through and clean their code of confusing parts.

