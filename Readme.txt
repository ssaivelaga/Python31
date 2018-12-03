Details of code files:
--------------------------
1. Main.py: This file contains the code of GUI interface to analyze the data of four companies, displays financial statements of different companies, plot graphs based on that data.It also compare and plot data of two companies.

2. prediction.py: This file contains the code to plots a predictive model for four individual companies as four different subplots.


DATASET Information
-------------------------

1. Comparison.xlsx
This excel file contains data for Financial Ratios, Comparison of returns of four companies. i.e NCR, CoBiz, Fiserv, CATHAY.


2. prediction.xlsx
This excel file contains data of four companies i.e NCR, CoBiz, Fiserv, CATHAY. The data includes year, Quarter, asserts,liabilities,and Income details

3. NCR Financial Data.xlsx/CoBiz Financial Data.xlsx/Fiserv, Inc Financial Data.xlsx/CATHAY Financial Data.xlsx
This excel files contains financial data of related companies .i.e Data for Result of operations, Financial Condition, Comprehensive Income,Consolidated Balance Sheets, Intangible assets.


The method used to create GUI
-------------------------------------------------
For making the GUI, we used the python ‘tkinter’ library. ‘tkinter’ is the basic library used for
creating GUI in python. The different wizards such as Entry field, Button, combobox etc. are the
classes defined in the tkinter library. For creating a table to be displayed in the separate window,
we used a number of Entry wizards where each Entry field is a single cell of the table. The Entry
fields are arranged using the Grid layout. We had settled the Window size and size for each
wizard manually.

Another libraries used:
-------------------------------------
Following is the list of libraries used in the program:
- ‘pandas’ – Used for creating dataframe from Excel files.
- ‘matplotlib’ – Used for visualizing the data i.e. for Plotting the graphs.