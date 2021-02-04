# VBA_Stock_Analysis
An analysis of a variety of stocks using Excel and VBA

## VBA Stock Analysis

Steve and his parents are looking to invest in green energy which they believe will be the energy of the future.  They are having a look at the stock price changes in comapnies that are developing green energy.  We have put together a list of stocks along with the daily volume and returns so that Steve and his family may make the most of their investment.

### Background

Steve has just graduated with a finance degree and is going to invest with his parents.  His parents believe that there will be more reliance on green energy in the future as fossil fuels get used up and are looking into a company called DAQO that produces green energy.  There are many different types of alternative energy sources such as hydro, solar, etc. and Steven would like to look into a few different stocks for diversification.  He has asked us to look into a variety of different stocks that he has chosen.

### Purpose

Steve's parents would like us to look into DAQO stock to see what kind of returns the company has been getting.  Steve would also like to diversify his stocks, so he has chosen a few other alternative energy stocks that he would also like us to look into.  The purpose of the analysis is to check the daily volume and returns on a handful of stocks that Steve has chosen in order to determine which are the best investments.

## Analysis

The vast majority of the green energy companies were doing a great deal better in 2017 than in 2018 experiencing a great deal more growth.  Additionally, we were able to make some improvements to the structure of our code to make it run faster making it more scalable for a larger analysis in the future.

#### 2017 Results

![2017 Stock Performance and Code Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### 2018 Results

![2018 Stock Performance And Code Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Stock Comparisons

#### 2017 Stock Performance

![2017 Stock Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/2017%20Results.png)

The vast majority of the green energy stocks were producing positive returns in 2017.  All but one of them produced positive returns.  The stock that produced negative returns was the TERP stock and the all the others produced positive returns.  There were four stocks, DQ, ENPH, FSLR, and SEDG, which produced 100% returns in 2017 making them the top performers.  DQ was the best performer at 199.4% and SEDG a close second at 184.5%. Companies such as SPWR and FSLR had a great number of daily volumes as well.

#### 2018 Stock Performance

![2018 Stock Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/2018%20Results.png)

The following year, 2018, did not offer great returns for investers.  There were actually only two companies that were in the positive which turned out to be ENPH and RUN, both of which were above 80% returns.  The rest of the stocks were in the negative DQ and JKS producing the worst returns both at less than -60%.  Despite this, nearly all of the companies were still at daily volumes of hundreds of millions except one which was AY.

### Code Optimizations

In the first iteration of the code we opted to write the information to the file within the for loop which takes a large amount of calculation.  We were able to change this to storing the information in an array and then writing all the information to the file at once which means that the file does not need to be accessed repeated, but only once.  This improved the code perfmance by about half a second even with only a dozen stocks.  This makes the code a great deal more scalable and would make a big difference if we were to run the code on thousands of stocks

#### Original Code Performance

![2017 Original Code Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/2017%20(orig).png)
![2018 Original Code Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/2018%20(orig).png)

#### Refractored Code Performance

![2017 Refractored Code Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/2017%20(Refr).png)
![2018 Refractored Code Performance](https://github.com/ForTheGold/VBA_Stock_Analysis/blob/main/Resources/2018%20(Refr).png)

## Summary

The goal of code refactoring is to opitimize the code and remove any bugs, inefficiencies, or other similar problems, but can the process can take time and money, cost resources, and introduce new bugs.

### Pros of Code Refractoring

The goal of code refactoring is to make the code better, which can mean making it more readable, optimizing performance, improving maintainability, etc.  Code should be refactored when there are clearly several repeats in the code that can be separated out into functions or loops that will optimize performance and readability.  Additionally, the code may contain bugs that can be removed during refactoring.  These are the pros of code refractoring.  The code can be made more maintainable, faster, more scalable, readable, etc.

### Cons of Code Refractoring

There are time when it is better to leave the code alone.  For instance, we do not make any large changes to the code when we are nearing the deadline for launch.  The programmer who is refactoring the code may end up introducing bugs in the code unintentionally that push the launch date of the code past the anticipated deadline.  Additionally, there may be times when making optimizations within the code does not make a great deal of difference in terms of performance.  Alternatively, the code may be so poorly written that it is better to start from scratch rather than make any attempt at optimizing the code as it will be faster to simply start over.  The cons of refactoring code is that can take a great deal of time, cost money, introduce new bugs, etc.

### Pros and Cons of Optimizing Our Code

Our original code had some serious problems with it.  It takes a great deal of processing power to open a file, write to it, and close it to do further calculations over and over.  This type of operation is much better done all at once.  This was an instance where very clearly we can make a large improvement without a great deal of time spent on fixing the code.  The cons of this is that it takes time.  If we were planning on running this analysis on only twelve stocks as we have now, then there is really no reason to make the code half a second faster as it makes no difference to the user.  This optimization is really only necessary if we are planning to scale this analysis to a much larger number of stocks.
