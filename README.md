# VBA-challenge
Submission for Module 2 VBA challenge.

Most of the code was derived directly from examples and learnings completed in class. The following pieces of code were realized through other methods:

- Dim stock_volume As LongLong (line 19): When the code first halted within the first few lines and output an error (runtime error 6), I Googled the error to discover that the variable type I had originally selected (Long) was insufficient, and needed LongLong instead.
- Pull each new ticker (lines 31-44): At first, I set up my code to pull each row of outputs at a time, starting with the ticker, followed immediately by that ticker's yearly_change, percent_change, and stock_volume. However, I ran into errors later in my coding. I worked with a classmate, Cameron Lee, who showed me that he pulled each criteria (e.g. ticker, yearly_change, percent_change, and stock_volume) one at a time as columns first, rather than rows. I copied his approach in pulling each ticker first, then used my approach in pulling the yearly_change and percent_change as rows for each ticker before moving on to the next ticker. I copied Cameron's approach again to pull the stock_volumes.
- Pull each ticker's year_open figure (lines 52-53): This was particularly difficult for me, as at first I could not figure out how to pull the year_open figure. I Googled the approach, and found the following code to pull the year_open figure if the preceding ticker was different. I then figured out that I needed to place this code first in the IF statement to avoid the formula only pulling the year_close figure in the yearly_change calculation.
	        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
	            year_open = ws.Cells(i, 3).Value
- WorksheetFunctions found via Google (lines 121-141): Since we did not cover this coding in class, but our instructor provided a tip that we'd need a WorksheetFunction, I Googled this approach pretty heavily to fully understand how it functions. I found that a site called wallstreetmojo.com was the most helpful, as it provided the most basic explanation. 
