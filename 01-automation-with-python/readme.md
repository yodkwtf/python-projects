# Automation with Python

A program where we are using python to process spreadsheets. It saves us an ample amount of time and processes everything in seconds which normally would take hours and days.

## Question Statement

Suppose there was an error in the transactions sheet and the prices inside need to be dropped by 10% to fix it. Write a program to automate the reduction of price by 10% for every product. Also add a chart based on it.

## Steps Taken to Solve

- Import `openpyxl` package and classes to create charts from it
- Create a reusable function that takes a file name as an argument

#### Fixing Prices

- Load the spreadsheet and select the required sheet from it
- Loop over the sheet cells to find and correct the faulty values
- Save the corrected values in new cells

#### Creating Charts

- Select the chart parameters/range using `Reference` class
- Create the chart and add data to it based on the parameters
- Add the chart to the sheet
- Finally, save the updated spreadsheet!
