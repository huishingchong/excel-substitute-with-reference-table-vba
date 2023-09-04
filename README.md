# excel-substitute-with-reference-table-vba
## Introduction
Excel VBA that substitutes all instances of cell content within a range of excel sheet, if the instances exists in the reference table supplied by the user, with corresponding values to substitute them with.

It will prompt user for two inputs:
- to select a range of cells in an excel sheet where content needs to be substituted
- supply a _local_ excel file containing the reference table for the substitution

## Requirements
The reference excel file will need to have a table to lookup on instances to substitute and what value to substitute instances with.
It is assumed that the reference table in the supplied reference excel file adhere the following requirements:
- Reference table is a Table on Excel
- Reference table is in a worksheet named "Sheet1" of the Excel file supplied
- Reference table is named "Table1" (this is also the default by Excel)
- Reference table contains _two_ columns
  - Column A: containing key (string to be replaced for lookup)
  - Column B: containing their corresponding value to replace (string to replace key with)

## Instructions to use VBA
If you're not familiar with running VBA, you can follow the steps below.
1. Start with opening your Excel file where you want to run the vba on. (the Excel sheet you want to perform substitution on)
2. Check that you have a 'Developer' Tab in the ribbon. If you don't have a 'Developer' tab, follow [these instructions]([url](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45#:~:text=How%20to%20Get%20to%20the%20Developer%20Tab%20in,%2C%20select%20the%20Developer%20check%20box.%20See%20More.)https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45#:~:text=How%20to%20Get%20to%20the%20Developer%20Tab%20in,%2C%20select%20the%20Developer%20check%20box.%20See%20More.) to enable it.
3. In the 'Developer' tab, Click 'Visual Basic', a pop-up screen will appear
4. Click Insert > Module and a blank editor will appear
5. Copy the VBA code in the github repository to the editor
6. Click _Run_ (or the green arrow) - it will ask you which Macro to run, select the correct one, in this case, SubstituteMain

## Things to improve on
- VBA code currently works only if requirements detailed above are met, does not contain error handling.
