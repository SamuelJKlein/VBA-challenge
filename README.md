# VBA-challenge

VBA challenge submission

VBA_Script.rtf is the script I wrote 
Results.png is the screenshot of the results


The code starts by creating some variables that are used later when we loop through the cells. We define them first so that they don’t need to be declared every time the code loops through a row. 

Then we loop through all the items in the worksheets object. 
Inside each worksheet, we start by getting the number of non-empty cells in the 1st column. That gets saved as a variable called RowCount. Then we use the RowCount variable as the limit for another “For-Next” loop that goes through all the rows, one at a time. 
The first thing we do in each row is add the volume from that row to the TotalVol variable. We get the opening and closing numbers by comparing the current cell value to the value either above it or below it, respectively. These numbers are saved as the YearStart and YearEnd variables. 
In the same “if” statement where we define the YearEnd variable, we also display the required information about that company: the yearly change, which is calculated using the YearEnd and YearStart variables, and the total volume from the TotalVol variable. This is displayed using “Worksheets(1)”, so that the summary list is shown on the first sheet of the Excel file. 
We also increase the value of DisplayRow, which is used with the Cells() function so that each company’s information will be printed in consecutive rows, instead of overwriting the previous company. 
The information from the company is then compared to the values of some cells on row 17 of Worksheets(1), so that if the current company passes the saved value for the greatest increase/decrease/total volume, then the number will be adjusted appropriately. 
Finally, we reset TotalVol to 0, so we can start over for the next ticker. 

At the end of the script is the formatting and headers for “Worksheets(1)”, which is where we are displaying the summary table. Some of the columns are set to auto fit so that the text and numbers are readable, and some cells and rows are formatted as percentages. 
We also count the number of rows in the summary table and apply conditional formatting to all of the used rows in columns J and K except for row number one, which is the header. I wrote it like this because otherwise it was making the header green, and I didn’t want that. 
