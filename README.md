# Stock Analysis in VBA
An analysis of some stocks to see trends in VBA

---

## Overview
I was asked by my friend Steve to assist him in analyzing a dataset of a dozen stocks he has selected. The dataset contains the opening price, closing price, daily high price, adjusted closing price, and the volume for a given date. There is a daily entry for each of the 12 stocks spanning 2 years. The years are split into two separate sheets.

Our goal is to help Steve see the trend of the stock over the year to better help him understand if the stock is sound investment. I also was interested in seeing if there could be improvements in my code to reduce the runtime.

## Results
### Stock Results
Upon review, we can see that only two stocks that produced a positive return in 2017 produced a positive return in 2018. 

The 2017 returns:

![2017 returns](./resources/2017_Stock_Performance.png)

The 2018 returns:

![2018 returns](./resources/2018_Stock_Performance.png)

So based on these returns, I would only be able to recommend Steve to look more into ENPH, as it seems as if it is the only viable option over the time frame he wanted to analyze. 

### Code Performance Results
Looking at the performance of the code, as we can see below: 

2017 Refactored Code:

![2017 Refactored](./resources/VBA_Challenge_2017.png)

2017 Original Code:

![2017 Original](./resources/VBA_Challenge_2017_original_code.png)

2018 Refactored Code:

![2018 Refactored](./resources/VBA_Challenge_2018.png)

2018 Original Code:

![2018 Original](./resources/VBA_Challenge_2018_original_code.png)

We can see a large discrepancy in the speed that it was able to finish the calculations. The speed of the refactored code is much faster in its execution compared to the originally written code.

## Summary
### Refactoring Code Benefits
Many times code written initially is not the best in terms of run-time, and can be improved. Refactoring code can also help in seeing what can be changed using different code practices. These code practices can include making code class based, or creating functions (or subroutines in the instance of VBA) which are called multiple times to make the code cleaner. So overall it can be very beneficial to refactor code to see where improvements can be made or how to make the code cleaner.

### Original Code vs Refactored Code
In this situation, our original code is running much slower than the refactored code. This is solely attributed to the way we were looping through the data. In the original code we were reading the entire dataset for every stock compared to only reading the dataset once in the refactored version. In order to fix the original we added arrays to store the data that we were looking to calculate, compared to originally saving it in a variable and printing to the file one ticker at a time. This method increased our space complexity by a factor of t where t is the number of tickers that we are analyzing. This however is a worthwhile trade off as the time complexity is reduced greatly. From $O(n^2)$ in the original code to $O(n)$ in the refactored code based off of changing the structure of the for loop from being nested to using an array to save the data separately. 