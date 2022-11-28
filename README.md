# stock-analysis

## Overview of Project 

The purpose  of this project was to refactor code to perform an analysis on the Total Daily Volume of twelve stocks, from the years 2017 to 2018. The code was refactored to run more efficiently, accounting for the possibility of larger quantities of data needing to be analyzed with the same script. The code will also produce the total time it took to pull and analyze the selected data. The analysis was conducted to investigate the stocks' returns. The script was written in VBA to be used in Microsoft Excel.

## Results

### 2017_Non_Refactor
![2017_Non_Refactor](https://github.com/bikachuuuuuu/stock-analysis/blob/main/Resources/2017_Non_Refactor.PNG)

### 2017_Refactor
![2017_Refactor](https://github.com/bikachuuuuuu/stock-analysis/blob/main/Resources/2017_Refactor.PNG?raw=true)

As can be seen evidently through the photos above, the initial code took approximately 91 miliseconds in order to run the code that was produced through the course of Module 2. With various degrees of modifications to the code, allowing the script to run through the data given at a much quicker pace, we see a decrease in the amount of time it takes to run the specific code. This is seen as a net gain, both for the developer and for the client as well. Therefore, the refactored code would be more useful to use in this circumstance to quicker provide the requested service to the client.

This can also be deduced, by looking at the 2018 results as well.

###### 2018_Non_Refactor

![2018_Non_Refactor](https://github.com/bikachuuuuuu/stock-analysis/blob/main/Resources/2018_Non_Refactor.PNG?raw=true)

###### 2018_Refactor

![2018_Refactor](https://github.com/bikachuuuuuu/stock-analysis/blob/main/Resources/2018_Refactor.PNG?raw=true)

In this example as well, much like the 2017 results, it can be determined that the refactored code has improved over all functionality of the script and is running at an ideal speed of what one would desire when running data anaylze macros.

## Summary

One major advantage to refactoring the code is that the code will be more precise, having fewer steps to take, and therefore execute the script at a quicker rate. However, a disadvantage to that because the code has become more specific, there's s chance of running into bugs while refactoring the code. As a result, it could possibly take time to run iterations of the code and perform a de-bug.

These pros and cons were reflected in this demonstration, as we were able to see how sorting data before running such programs can drastically cut down on necessary code and allow for much more efficient runtimes. However, we were able able to see how quickly changing small things can snowball into runtime errors, as having to go back and correct your variable names can create de-bugging issues.
