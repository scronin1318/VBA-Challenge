# VBA-Challenge

## Overview of Project

The purpose of this analysis is to calculate the total daily volume and return rate for each of the twelve stock tickers for each year (2017 and 2018) as efficiently as possible. By running this analysis, we can ascertain which stocks are the most profitable (by looking at return rates), as well as which stocks are most liquid (by looking at total daily volume), which can ultimately help us decide which stocks are worth investing in. Specifically in regards to VBA, our goal is to refactor code so that it runs quicker and more efficiently.  

## Results

  ### Comparing 2017 and 2018 Stock Performance
  <img width="242" alt="Screen Shot 2021-05-08 at 3 37 00 PM" src="https://user-images.githubusercontent.com/82490011/117552799-409f9700-b013-11eb-9405-d82cff10374a.png">
  <img width="235" alt="Screen Shot 2021-05-08 at 3 38 40 PM" src="https://user-images.githubusercontent.com/82490011/117552843-7e042480-b013-11eb-9742-a3d7c947ac85.png">

  Loooking at the 2018 versus 2017 stock performances, it's clear that 2017 stocks performed much better than 2018's, as indicated by the positive return rates (eleven out of the twelve stocks have positive return rates). Looking at 2018, ten out of the eleven stocks have negative return rates. Calculating quick averages, 2018 stocks have an average return rate of -8.5%, whereas 2017 stocks have an average return rate of +67.3%, which further validates our conclusion. The only two stocks' performances that improved between 2017 and 2018 are RUN and TERP, going from a return rate of 5.5% to 84.0% and -7.2% to -5.0%, respectively. 


  ### Comparing Execution Time Between Original & Refactored Script
  Looking at the below screenshots, the original (un-refactored) script took ~0.32 seconds to run for 2017 and 2018
  
  <img width="256" alt="Un-refactored_2017" src="https://user-images.githubusercontent.com/82490011/117553804-28327b00-b019-11eb-91f5-a3168762a1ee.png">

<img width="258" alt="Un-refactored_2018" src="https://user-images.githubusercontent.com/82490011/117553807-2c5e9880-b019-11eb-90cd-dea2a22d9d47.png">

  ...whereas the refactored script took ~0.10 seconds to run.
  
  <img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/82490011/117553848-73e52480-b019-11eb-8c9d-64b4c5bc55fb.png">

<img width="258" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/82490011/117553852-79db0580-b019-11eb-9f0b-a2ad4387e2ae.png">
  
  Refactoring the script took ~0.22 seconds off of the run time, improving the speed by 69%! 

## Summary

  ### Advantages & Disadvantages of Refactoring Code
The advantage of refactoring code is that it increases efficiency. Instead of looping through our tickers each time when we want to calculate the volume or      find opening and closing prices, we can use an index to start with our first ticker, find all the necessary information we need, and then move onto the next ticker once we've done so. This allows for increased efficiency, and as a result, shorter script run-times. Similarly, it makes the code more streamlined and therefore easier to maintain. The disadvantages of refactoring code include the fact that it can be time-consuming, as well as the fact that it may introduce bugs into the code. 

  ### How do these pros and cons apply to refactoring the original VBA script?
The imporved efficiency can be seen most notably in our pre-existing if-statement used to determine a stock's ending price. With the refactored code, we can concurrently determine if it's time to move onto the next ticker or not, as shown in the below code. By not requiring a for-loop to loop through the tickers, and instead running the analysis one ticker at a time, our refactored code is more efficient. 




  <img width="596" alt="Screen Shot 2021-05-08 at 4 32 57 PM" src="https://user-images.githubusercontent.com/82490011/117554099-12be5080-b01b-11eb-866b-1d8fe19c8b62.png">
  
  
  
Conversely, in the original script, a stock's ending price was calculated within a for-loop that looped through the tickers. So not only did our computer have to loop through the rows in evaluating our data, but it had to loop through the tickers as well, creating a for-loop within a for-loop, decreasing efficiency.




<img width="596" alt="Screen Shot 2021-05-08 at 4 32 57 PM" src="https://user-images.githubusercontent.com/82490011/117554099-12be5080-b01b-11eb-866b-1d8fe19c8b62.png">  

While it's true that refactoring code can be time-consuming and can potentially incorporate bugs into our code, the drastically increased efficiency outweigh the cons. After all, we improved our computational speed by 69%! 
