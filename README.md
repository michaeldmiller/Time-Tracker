# Time-Tracker
This program is a time tracking assistant I wrote in June of 2020 to help keep better track of 
quarantine days which seemed to be slipping away at an alarming rate. 
It depends on the docx, openpyxl, numpy, and matplotlib external libraries for basic function.
Based on a text input file, day_summary.txt, it takes data about the categories of actions which the 
user reports having occurred during their day and produces two primary outputs: a horizontal 
stacked bar chart and a labeled pie chart. It then integrates these two charts together, saving 
them in a Word document. It then takes the raw data and stores it for later use in an excel file, 
fittingly named RawData.
It is designed to be refreshed on the first of every month, saving the user's information data locally
on Word documents Week1 through Week4, covering 28 days. If the month is longer than this, the user
will need to remove the filled Word documents from the main directory, and let the program start a new 
"Week1", which it can run until the end of the month. I usually then rename this Week5. The excel file 
can remain theoretically indefinitely, though performance slows if it is allowed to accrue too much data. 
4 months is a good target, I have found that 6 months of data gets to be intolerably slow.
