# excel-substitute-with-reference-table-vba
Excel VBA that substitutes all instances of cell content if the instances exists in the reference table supplied by the user, with corresponding values to substitute them with.

It will prompt user for two inputs:
- to select a range of cells where content needs to be substituted in an excel file
- supply a reference excel file containing the lookup table for the substitution

The reference excel file will need to have a table to lookup on instances to substitute and what value to substitute instances with.
It is assumed that the reference table in the supplied reference excel file adhere the following requirements:
- Reference table is in "Sheet1" of the Excel file
- Reference table is named "Table1" (this is default by Excel)
- Reference table is a Table on Excel
- Reference table only contain 2 columns
  - First column containing key (string to be replaced for lookup)
  - Second column containing their corresponding value to replace (string to replace key with)
