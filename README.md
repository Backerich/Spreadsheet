# Spreadsheet

## Overview
SPREADSHEET is a tool that automates finding according data in workbook B and replacing them in workbook A.
It is similar to the VLOOKUP Function in Excel.

## Example
**Workbook A:**

Zero | First Header | Second Header
---- | ------------ | -------------
First | Content from cell 1 | Content from cell 2
Second | Content in the first column | Content in the second column

**Workbook B:**

Zero | First Header | Second Header 
---- | ------------ | -------------
First | Content from cell 1 = "Changed" | Content from cell 2 
Second | Content in the first column = "Changed too" | Content in the second column

**Description:**

SPREADSHEET finds the data which intersects. Then it searches in workbook B for the equalsign ('=').
When this displays, the data after the equalsign should replace the relating string (infront of the equalsign) in workbook A.
(in this example 'changed').
After the programm was executed the cell "First Header"/"1" should now say 'Changed' in workbook C.
The programm copies workbook A into a new workbook C and adds the results of the programm.

**Output Workbook C:**

**Workbook A:**

Zero | First Header | Second Header
---- | ------------ | -------------
First | Changed | Content from cell 2
Second | Changed too | Content in the second column
