# excel-vba-data-cleaning

Problem Statement

This assignment starts with a bunch of orders that have recently been placed for your publishing company.  Each week or so you receive an Excel file with all the orders, which include the Date Ordered, Store, Title of the book, and Quantity.  The starter file looks like this:
![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/9872de94-12c1-4c1d-afbb-1c48cfd83988)


Goal #1: Data cleaning and identification of incomplete orders

Your first goal is to clean up the data.  First, you want to remove any formatting in the table.  For example, you want to eliminate the red font color in rows 9 and 34 and the bold formatting in rows 9 and 27.  You should set up your subroutine such that it will adapt to *any* order sheet – formatting could be in any cell, so make sure you are eliminating any formatting in all cells.  Furthermore, the number of orders in subsequent order sheets could be different; therefore, your sub needs to adapt to the number of orders and automatically adjust.

Another aspect of data cleaning is to remove entirely any rows that have a missing Store.  If the Title or Quantity is missing, it’s fairly easy to follow up with the Store to let them know that their order was incomplete – in fact, this is one of the main aspects of this assignment, as you’ll see below.  However, if the Store is missing, then it’s impossible to follow up, so any rows with missing (blank) Store should be eliminated entirely as part of the data cleaning process.

In the starter file, you should place the code for the above cleaning process into the first part of the FormatAndIncompleteOrders sub (which is linked to the “FORMAT & GENERATE INCOMPLETE ORDERS REPORT”).

Within the FormatAndIncompleteOrders sub should also be code that will generate a report of any incomplete orders.  This means that any rows that have either the Title or Quantity missing (blank) should be copied over to the “Incomplete Orders” sheet.  These rows would also be eliminated from the original data on the “New Orders” sheet.

IMPORTANT: 

Depending upon how you do things, you may end up inadvertently changing the format of the Date Ordered column to general (it’ll just be a serial number, like 44805).  If this happens (it probably will), then you can record a macro to reformat the dates to Short Date format and implement this code into your sub.

Make sure that you are NOT deleting the column headers (“Date Ordered”, “Store”, etc.)!

The Reset sub is only used to reset the ORIGINAL data.  You should NOT call (refer to) the Reset sub from within your FormatAndIncompleteOrders sub!  If you are having issues with the grading file, you might check to make sure that you are not calling the Reset sub in your FormatAndIncompleteOrders sub.

The result of running the FormatAndIncompleteOrders sub on the data in the starter file would result in the following:
![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/60a3831a-ef33-4c03-a3fd-d4302b844e77)


Furthermore, at this point the “Incomplete Orders” sheet would appear as follows:

![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/643cdbec-7fe4-4ccf-bf28-78106da3bf4a)


The RESET button at any point will reset the original data (available on the “Original Data” sheet) in the “New Orders” sheet by running the Reset sub.  Do not modify or remove the Reset sub as the grader file will run the Reset sub to reset the data when it grades your work.

HINTS

To identify rows that have blanks, you can filter by blanks (record macro if you need to, or you can just have as the criteria “” (empty quotations).

To move rows with blank Title or Quantity, you can copy/paste into the “Incomplete Orders” sheet and then delete them from the “New Orders” sheet.

Make sure that you are counting the number of orders so that your sub adapts to the number of orders – the grader file will test a different set of data with greater or fewer number of orders!

    

Goal #2: Store-specific report

The second main objective of this assignment is to generate a Store report from the Store that’s selected from the drop-down menu in cell H10.  After a Store is selected in the drop-down list, the user can press the GENERATE STORE REPORT button, which runs the Report sub, the data will be filtered by column B, and only those rows that match the Store in cell H10 would be selected/copied and pasted over to the “Report” sheet.  Note that you are NOT removing the rows from the original data in the “New Orders” sheet but are just copying over those orders that correspond to the Store in cell H10.

For example, if the user selects Bob’s Books from cell H10 and runs the GENERATE STORE REPORT button, the “Report” tab would appear as follows:

![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/ccdb4182-1019-4aa8-88d1-32ff17b46be8)


Like Goal #1 above, it is important that your code adapts to different data.  For example, if we ran the sub on a worksheet that had a different number of rows, then it should work fine.
