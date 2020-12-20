# vandcamp_Mod2_VBAchallenge
Analysis of Challenge 2

## Overview
In this challenge, the analyst is tasked with creating dynamic VBA macros that will aid the client, Steve, in visualizing data for 12 different stocks in a specific market year. While the client has 12 specific ticker symbols to watch in their portfolio, the idea is to create a tool that would allow them to visualize the performance of each security after choosing a specific year. 

## Stock Performance
For Steve, 2017 was a much better year for his portfolio, since each of these stocks gained an average of 67% in the market year. That said, the sum loss of 2018 does not overshadow 2017's gains. Hopefullyl his portfolio hasn't taken too much of a hit into the treacherous economy of 2020.

# VBA Subroutine Analysis
In this challenge, the analyst begins with pieces of code created to address specific requests one at a time. By refactoring these requests, the end result is a subroutine that handles all of these requests in batches rather than as an individual function. Let's begin with the biggest change in functionality: data storage.

## Data Storage
The new code utilizes arrays in order to store the analysis of each stock, rather than calculating performance, outputting the result, then going back to the data. This means that the original program requires a Worksheets(x).Activate function at the beginning and end of each loop. On the contrast, this refactor stores all of the math and value lookups in arrays that we can call back to once we're ready to display the sum of the analysis. The biggest step in streamlining this process is the initialization of tickerIndex, as seen below. 

![Let the computer do the thinking!](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/data_streamline.png)

By creating tickerIndex as an integer within the macro, we can utilize that phrase as a variable within arrays we already have created. See how at the bottom of the image that tickerIndex is defined as our loop counter "j". As the computer goes through analysis of each ticker in the following If/Then loops, it stores the results of each under the loop number, or, tickerIndex. Then, once the computer finds all the values our client wants, the twelve data arrays come back to us with a simple output For loop, below.

![Show us the money](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/output_loop.png)

### Illustrative Add-Ons
To make the ultimate analysis a little bit more visually stimulating, we let the total volume breathe in its column with an autofit function. Similarly, I added some italics to the ticker symbols, and a biggest winner/loser tracker below our table. 
 
The biggest change to the base subroutine when it comes to formatting was a loop practice that colors the return column based on a positive/negative test. I decided not to bother with this color coding of the "Highest/Lowest" lookup below the output table, since the corresponding colors are a given, and less useful for the immediate gratification of seeing which stock won and which lost the year. 

![Crude, yet effective](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/format_practice.png)

I also added a clear button for the client that eliminates the previously enacted analysis. Due to the nature of the original subroutine's checks, it's not necessary to eliminate the formatting, but I figured it was good practice for my skills. On that 

## Speed Boost
This refactoring of data storage is a big change to the base code's order of operations, so it should accelerate the calculations by quite a bit. Below is screencaps of the base subroutine's timer.

![Base Code, 2017](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/VBA_lesson_runtime.png)

![Base Code, 2018](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/VBA_lesson_runtime2018.png)

Alas, we're only able to shave 0.2 seconds from our client's time by making this refactored macro (see below) but at least it has the capability to analyze another year's stocks, so long as a new worksheet is created with a dump of the year's stock data.

![2017 speed](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/VBA_challenge_2017.png)

![2018 Speed](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/VBA_challenge_2018.png)


### Shortcomings
The biggest shortcoming for the end-user is that any additional stocks added to the portfolio would need to be manually inserted into the ticker array, and each output array would need to be increased in size by however many more stocks we're watching. I have a hunch that I could create a dummy worksheet the end-user could fill in with ticker symbols that would allow the program to count and create an array of how many stocks to look out for. However, I am tired of working with this particular worksheet and am choosing to move on. I'll save this ambition for a client in the real world.

The only other big shortcoming is in the form of UX. I think it's lame to let the user end up with a runtime error should their finger slip and they don't choose 2017 or 2018 to analyze. While I did make a quick error box that works in the bounds of the challenge, it no longer makes the macro dynamic enough to function should more years of data be added to the workbook as a whole, so I ultimately commented it out. I lack the patience to make the program constantly scan for an error code 9, and that is my shortcoming.

![Polite, yet firm](https://github.com/davidandcarr/vandcamp_Mod2_VBAchallenge/blob/main/Resources/error_protocol.png)