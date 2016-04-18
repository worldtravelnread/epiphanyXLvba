
# epiphanyXLvba

## Excel VBA code to clean BCMS data

While writing this code, I realized that I loved coding and working with data.
This epiphany led to my decision to transition to a career in data science.

The reports from BCMS, a contract management workflow system, had thousands of
duplicate records and many fields had zero-length strings (i.e., ""). 

I wrote the code to help troubleshoot the reports. This is not hard coded to
work only with the BCMS data. It can be used on any data.

The first module creates a new sheet, "notes", to record information about 
the data and another sheet to record information about the data source.

Excel has a built-in function to remove duplicates from data, but it does not
give you the option to put the duplicates in a new worksheet, which might be
helpful in troubleshooting the data.

The second group of modules copies the sheet with the raw data and identifies
the duplicate records. Then a pivot table is created, from which a detail sheet
of only the duplicates is made. The code identifies the first occurrence of 
each duplicate record. The number of duplicate records is entered into the
"notes" sheet.

The third code module identifies which fields have zero-length strings. It does
this for all records and records the number of zero-length entries in each
field in the "notes" sheet.

The fourth code module creates a named range of all the data in a sheet. This
is to make VLOOKUP formulas easier to read.







=======
# epiphanyXLvba
Excel VBA code to clean BCMS data
>>>>>>> c3e97367fed2e2a2d03091fc0e71ab2023a8cf58
