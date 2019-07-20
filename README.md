# BARFTool
A tool used to quickly search an Excel file for names and emails of employees.

BioMeds, ATOMSs, RTOMS, and FVPs (BARF) are the titles of a series of ascending technical people in our company. An Excel workbook of nearly 20 sheets is distributed monthly detailing who reports to whom and where the lowest level of employees are located. This workbook is a real pain to search for information beacuse of its multiple sheets and redundancy. 
So, I created a little python script that searched through the sheets using openpyxl. It starts by finding the BioMed associated with a clinic, searching through the hierarchy to find their boss, searching through again to find the boss's boss, etc. until it found all the technical contacts associated with a site.
It had to handle the inconsistencies in how the data was entered on different sheets. I'm afraid I cannot provide a copy of the workbook, itself, because I am not certain if the employee data is confidential or not.
