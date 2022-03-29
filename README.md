# Stock Analysis using VBA in Excel

## Overview of Project

The purpose of this project was to help Steve be able to expand his stock analysis beyond the twelve stocks he was originally looking into for his parents, to include the entire stock market. In order to make that possible we had to refactor the original VBA code we created to make it more efficient so that it could potentially handle a much larger dataset (entire stock market), thus enabling Steve to make additional stock recommendations to his parents based on the expanded results.   

## Results

By refactoring the VBA macro my computer was able to run the analysis much faster, cutting the time from an average of ~8 seconds down to less than 1 second.  Examples of the time elapsed time for both 2017 and 2018 showcase this improvement in the macros efficiency

![2017](https://github.com/NRFlood/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018](https://github.com/NRFlood/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

Creating the tickerIndex variable allowed me to adjust the original code to remove one IF THEN statement that was calculating the total volume of each ticker, and simplify that part of the code to be a straight forward function (*tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value*). The tickerIndex also enabled us to remove a nested FOR LOOP statement where the macro was searching to find the first and last price of each stock, and then performing the output calculations within the same FOR LOOP. With the tickerIndex variable I wrote a second FOR LOOP that determined the start and end price of each ticker, and then increased the tickerIndex to the next ticker when complete.  From there I was able to complete all output calculations with one last FOR LOOP that leveraged the tickerIndex variable to access the correct index in each array.  These changes resulted in the increased efficiency referenced above.      

## Summary
### What are the advantages or disadvantages of refactoring code?
The biggest advantage of refactoring code is the ability to make it more efficient, and a more manageable file size.  The biggest disadvantage I can imagine would be working through any bugs as you go.  Editing an existing code that already works means you need to know every single part of the code that needs to be adjusted in order for it to continue working. With larger codes that could be a daunting task I would imagine as troubleshooting your way through the code if it doesn't work may take a great deal of time. 

### How do these pros and cons apply to refactoring the original VBA script?
In this case refactoring the code greatly improved Steve's ability to analyze more data in a shorter amount of time.  The cons I found were in editing the code correctly to continue capturing the relevant information. I struggled to find all of the parts of the code the need adjusting in order for the macro to run faster, but the exercise is worth the time if it means that a repeated process can now be completed more efficiently going forward.  
 

