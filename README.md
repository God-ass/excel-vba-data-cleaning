# excel-vba-data-cleaning

## Task 1: Publishing Company 

Problem Statement

This assignment starts with a bunch of orders that have recently been placed for your publishing company.  Each week or so you receive an Excel file with all the orders, which include the Date Ordered, Store, Title of the book, and Quantity.  The starter file looks like this:
![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/9872de94-12c1-4c1d-afbb-1c48cfd83988)


Goal #1: Data Cleaning and Incomplete Orders Identification

Data Formatting: Remove all cell formatting in the order sheet. The subroutine should be adaptable to any order sheet, regardless of the number of orders or cell formatting.

Missing Store Entries: Delete rows with missing Store entries. If Title or Quantity is missing, we can follow up with the Store. However, if the Store is missing, we cannot follow up, so these rows should be removed.

Incomplete Orders Report: Generate a report of incomplete orders (missing Title or Quantity) in the “Incomplete Orders” sheet. These rows should also be removed from the original data on the “New Orders” sheet.

Place the code for these processes in the first part of the FormatAndIncompleteOrders subroutine, which is linked to the “FORMAT & GENERATE INCOMPLETE ORDERS REPORT” button.

IMPORTANT: 

Make sure that you are NOT deleting the column headers (“Date Ordered”, “Store”, etc.)!

The Reset sub is only used to reset the ORIGINAL data.  You should NOT call (refer to) the Reset sub from within your FormatAndIncompleteOrders sub!

The result of running the FormatAndIncompleteOrders sub on the data in the starter file would result in the following:
![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/60a3831a-ef33-4c03-a3fd-d4302b844e77)


Furthermore, at this point the “Incomplete Orders” sheet would appear as follows:

![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/643cdbec-7fe4-4ccf-bf28-78106da3bf4a)


The RESET button at any point will reset the original data (available on the “Original Data” sheet) in the “New Orders” sheet by running the Reset sub. 
    

Goal #2: Generate Store-Specific Report

This goal involves creating a report for a specific store selected from the drop-down menu in cell H10.

Store Selection: Choose a store from the drop-down list in cell H10.

Generate Report: Click the “GENERATE STORE REPORT” button to run the Report subroutine.

Data Filtering: The data will be filtered based on column B to match the selected store in cell H10.

Report Creation: Rows matching the selected store are copied to the “Report” sheet. Note that these rows are not removed from the original data in the “New Orders” sheet, but simply copied over.

This process allows you to generate a report specific to the selected store.

For example, if the user selects Bob’s Books from cell H10 and runs the GENERATE STORE REPORT button, the “Report” tab would appear as follows:

![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/ccdb4182-1019-4aa8-88d1-32ff17b46be8)


Like Goal #1 above, it is important that your code adapts to different data.  For example, if we ran the sub on a worksheet that had a different number of rows, then it should work fine.

---

## Task 2: Extracting email addresses from mixed string formats

Problem Statement

Create a function that can identify and extract email addresses from unstructured text data.

There are 4 different formats on the worksheet, and we wish to extract only the email address:
![image](https://github.com/God-ass/excel-vba-data-cleaning/assets/92200827/b87edb0c-845e-4ce8-a72c-9e468414a1f1)





