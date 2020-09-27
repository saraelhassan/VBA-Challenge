# VBA-Challenge
A thorough analysis of stock market data for the years 2014-2016 was ran through using VBA scripts in Excel.
A ticker symbol was populated from the data provided over the years to show progressive or regressive growth.
Color coded the yearly change from opening price of the given year to the closing price at the end of that year, Green for growth and red for decrease.
Differentiated the precentage change between opening price at the begining of the given year to the closing price at the end of that year.
Provided the overall stock volume for each respectful year, with appointed greatest increase and greatest decrease of those years.


'declare the worksheets
 'Label Column Headers and Tables
  'Declare variables and set counter to default amounts
  'Determine value of the last row by finding the last non-blank cell in column A
  'Loop through rows
  'Add values to Total Ticker Volume
  'Check if the next row has the same ticker name as the previous row
  'Set Ticker Name for the first column
   'Print Ticker Name in Summary Table at Column I
   'Print Total Ticker Volume in Summary Table at Column L
    'Reset Total Ticker Volume
    'Set Yearly Open Price
    'Set Yearly Close Price
     'Set Value of Yearly Change
     'Change format of Column J to Accounting with "$"
    'Determine Percent Change, if Yearly Open Price is 0, then Percent Change is 0
    'Otherwise, set Percent Change to Yearly Change divided by Yearly Open Price
     'Print Percent Change to Column K
     'Change format of Column K to Percentage with "%" and to the hundredths decimal place
      'Conditional Formatting, if value is Positive, fill cell with Green
      'Conditional Formatting, if value is Negative, fill cell with Red
      'Add 1 to Summary Table Row
      'Set Previous Amount
      'Go to next row
      'Determine value of the last row by finding the last non-blank cell in column K
      'Loop through rows for final result table
      'Determine Greatest % Increase
       'Determine Greatest % Decrease
        'Determine Greatest Total Volume
      'Change format of Q2 and Q3 to Percentage with "%" and to the hundredths decimal place  
      end formula
      
