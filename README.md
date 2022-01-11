# Challenge 2-VBA Stock Analysis

### 1: Overview of Project

The purpose of this analysis was to learns the basics of Microsoft VBA and begin our coding journey. In doing this, we helped our friend Steve better understand how well a list of stocks performed in years 2017 & 2018, so he can give proper guidance to his parents on which stocks they should invest in. In this project, we examined a list of 12 stocks and what their total daily volume was for that year, as well as the return. Doing this gave us a better understanding of how well the stocks were performing, and if they were worth investing in.



### 2: Analysis

In 2017 all the stocks we analyzed did really well, with the exception of TERP coming in at -7.2% for the year. Otherwise, the rest of the stocks were positive, with DQ being at a positive 199.4% for the year, it seems they did well for themselves! 2018 is a different story though; almost every stock was negative for the year with the exception of ENPH and RUN. I am no expert in stock analysis or trading on the stock market, but it would seem ENPH and RUN would be the best stock options to invest in. Based on the data from the past two years it seems they are continuously growing and whatever money invested would do the same. Contrary to that, VSLR would be a good option to invest in, though it did drop in 2018, it was a very small amount, and you would not have lost nearly as much money as if you invested in DQ or JKS, which both dropped over 60%.

In regards to refactoring my code; the refactored code ran much faster than my original script. Below I have attached pictures of my original script, and the refactored code for 2017 and 2018.

Run time for original script

![Original VBA Script](https://user-images.githubusercontent.com/95730890/148866597-5090e869-3038-4168-b936-f4441904fb26.PNG)

Run time for 2017

<img width="400" alt="VBA_Challenge_2017-3" src="https://user-images.githubusercontent.com/95730890/148808167-3ea43776-4599-4951-8821-070e602b47ac.PNG">

Run time for 2018

<img width="400" alt="VBA_Challenge_2018-1" src="https://user-images.githubusercontent.com/95730890/148808292-025d2503-2ec1-42fe-8feb-f9d24f20e7d9.PNG">


### 3: Summary

In summary refacoring the code made it much faster than my original code. My original code ran in little over 2.0 seconds. But my refactored code ran in 0.4 seconds and 0.085 seconds. While this doesn't seem like much of a different, the refactored code ran at least twice as fast. When in a professional role, when we're dealing with exceptionally large files of code, this would most likely make a huge difference.

Also, refactoring the code made introduced me to some of the mistakes I made when I wrote my original script. It got my script to run the correct numbers with the correct formatting, but I had to add more lines of code to do this. So it was a great help to see how it could be done more effciently.
