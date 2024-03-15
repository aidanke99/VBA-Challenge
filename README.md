# VBA-Challenge

In this challenge, we will review multiple years of stock performance for multiole tickers. The goal is to calculate the total growth and decline of a ticker throughout a year, as a percentage, and also calculate that tickers total volume for the year. This script must be able to loop through all three of the years.

Method

to begin the project, I started by creating headers in a separate column, and calulated the yearly change of each ticker, and calculated this change as a percentage. To do this, I needed to know when the ticker symbol changed, so the loop would know when to reset, and move on to the next ticker, and recalculate

From there, I calculated the total volume of each ticker, by adding up the volume until the ticker symbol changed, and reset back to zero to recount for the next ticker.

Once total volume and percentage gained/lossed was calculated, I created a separate column to search for the ticker with the biggest gain for the year, the biggest loss, and the largest total volume. This could be done by searching for the maximum and minimum values in that column, and moving that value to the designated column.

Once I was able to run the code for the first year, It was simple to switch between sheets and calculate this information between all three years.
