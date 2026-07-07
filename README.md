# Excel UDF Add-In (AKHudfC)
This is a basic Excel add-in written with Csharp / ExcelDNA with the intent the assistance of the preparation of engineering calculations. It contains 17 User Defined Functions (UDFs). A few of the functions are not working as fully intended yet (see _docs folder for more information).

## INSTALLATION
1. After downloading and saving the XLL file, open Excel. 
2. Click File --> Options. This will open the "Excel Options" dialog box.
3. On the lefthand menu, click "Add-ins". Near the bottom, next to the "Manage [Excel Add-ins]" label, click "Go...". This will open the "Add-ins" dialog.
4. Click "Browse..." and locate / select the XLLL file.

The add-in inludes a ribbon tab with help and information regarding the use and capabilities of the add-in. Once loaded in the Excel environment, the tab should be visible.

<p><img width="765" height="127" alt="image" src="https://github.com/user-attachments/assets/f8fbf50d-eab8-409a-a865-e89859c15acc" />
</p>

## FUNCTIONS
All functions are prepended by ""C_"" for ease of location and input into cells.

<p><img width="401" height="203" alt="image" src="https://github.com/user-attachments/assets/f872780c-e121-491d-bbb4-a8c1b26e7ca6" />
</p>

| # | Function | Description |
|---:|----------|-------------|
| 1	| C_AbsMax	| Returns the absolute maximum (farthest from zero) of the entered numbers/references.	|				
| 2	| C_AbsMin	| Returns the absolute minimum (closest from zero) of the entered numbers/references.	|				
| 3	| C_BiggerNumber | Function to track the milliseconds required to count up to a number.	|	
| 4	| C_ChangeCase | Changes the case (upper, lower, sentence) of a string.	|
| 5	| C_ClampIt | Returns the specified value if it falls within the specified range, otherwise returns a limiting value.	|
| 6	| C_CountUnique | Counts the number of unique cell values in a range.	|
| 7	| C_FootToMix | Converts a number (assumed to be feet) to a mixed dimension string (ft-in).	|
| 8	| C_Formula | Similar to FORMULATEXT function, returns string of referenced cell's equation, showing values. |
| 9	| C_FormulaEq | Wrapper for C_Formula to prepend an equals sign and an optional label/unit. |
| 10 | C_InchToMix | Converts a number (assumed to be inch) to a mixed dimension string (ft-in).	|
| 11 | C_Indirect | Similar to INDIRECT function, but could access closed workbooks. |
| 12 | C_IsBetween | Returns a true/false boolean, indicating if a value falls within a specified range. |
| 13 | C_Linterp | Performs a linear interpolation, based on two known points (X, Y). |
| 14 | C_MixToInch | Converts a mixed unit (ft-in) string to a decimal number of inches. |					
| 15 | C_Mmatch | Returns the matching row number of a table, based on multiple criteria. |
| 16 | C_RevStr | Reverses the order of a specified string. |
| 17 | C_SheetName | Returns the worksheet name of the active cell. |
| 18 | C_WorkbookName | Returns the workbook name of the active cell. |
