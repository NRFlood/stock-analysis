# Stock Analysis using VBA in Excel

## Overview of Project

The purpose of this project was to help Steve be able to expand his stock analysis beyond the twelve stocks he was originally looking into for his parents, to include the entire stock market. In order to make that possible we had to refactor the orignal VBA code we created to make it more efficient so that it could potentially handle a much larger dataset (entire stock market), thus enabling Steve to make additioanl stock recommendations to his parents based on the expanded results.   

## Results

By refactoring the VBA macro my computer was able to run the analysis much faster, cutting the time from an average of ~8 seconds down to less than 1 second.  Examples of the time elapsed time for both 2017 and 2018 showcase this improvement in the macros efficiency (INSERT IMAGES BELOW).

Creating the tickerIndex variable allowed me to adjust the orignal code to remove one IF THEN statement that was calculating the total volume of each ticker, and simplify that part of the code to be a straight forward function (*tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value*). The tickerIndex also enabled us to remove a nested FOR LOOP statement where the macro was searching to find the first and last price of each stock, and then performing the output calculations within the same FOR LOOP. With the tickerIndex variable I wrote a second FOR LOOP that determined the start and end price of each ticker, and then increased the tickerIndex to the next ticker when complete.  From there I was able to complete all output calcautions with one last FOR LOOP that leveraged the tickerIndex variable to access the correct index in each array.     

## Summary

Refactoring the VBA macro...

