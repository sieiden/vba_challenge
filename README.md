# vba_challenge
Homework 2 for NW Bootcamp

This repo my vba code for the Wall Street stock assignment and three screen shots of my analysis (one for each worksheet). The second half of the second worksheet and the final worksheet (screen shot 3) are blank because my code throws a divide by zero error on line 38 (find the percent_change) for the ticker PLNT on the second worksheet that I could not figure out how to debug. I decided to submit the assignment as is so that I can move on to the Python assignments and keep up with the rest of the class.

The code loops through each of the rows in column A (ticker) and finds each unique ticker value in the list. A new table is created that lists each of the individual ticker symbols and their corresponding yearly change (ending value - opening value), yearly change percentage (yearly change/opening value), and total stock volume. The yearly change column is the conditionally formated so that negative values are red and positive values (including zero) are green. The Sub Greatest then finds the greatest increase, greatest decrease, and greatest total volume from the table of unique ticker values. Since the code throws an error in the middle of the second sheet, those values may or may not be correct and the third sheet doesn't list those values since the unique table was never created.

Google Drive Link to Excel File (Full Data Set) https://drive.google.com/file/d/1gHrCf8yYiTOW-x9UOsX3sn_8DV-9pA-J/view?usp=sharing
Google Drive Link to Test Excel File (Test Data Set)https://drive.google.com/file/d/1e-DjFkMwGhI5cIlQXQNqNuEBNwZj8lk7/view?usp=sharing
