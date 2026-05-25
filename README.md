This is a basic Excel add-in written with Csharp / ExcelDNA with the intent the assistance of the preparation of engineering calculations.

This add-in contains 17 UDFs (not all of which are working properly yet).

All functions are prepended by "C_" for ease of location and input into cells.


C_AbsMax		Returns the absolute maximum (farthest from zero) of the entered numbers/references.					
C_AbsMin		Returns the absolute minimum (closest from zero) of the entered numbers/references.					
C_BiggerNumber	Function to track the milliseconds required to count up to a number.					
C_ChangeCase	Changes the case (upper, lower, sentence) of a string.					
C_ClampIt		Returns the specified value if it falls within the specified range, otherwise returns a limiting value.					
C_CountUnique	Counts the number of unique cell values in a range.					
C_FootToMix		Converts a number (assumed to be feet) to a mixed dimension string (ft-in).					
C_Formula		Similar to FORMULATEXT function, returns string of referenced cell's equation, showing values.					
C_InchToMix		Converts a number (assumed to be inch) to a mixed dimension string (ft-in).					
C_Indirect		Similar to INDIRECT function, but could access closed workbooks.					
C_IsBetween		Returns a true/false boolean, indicating if a value falls within a specified range.					
C_Linterp		Performs a linear interpolation, based on two known points (X, Y).					
C_MixToInch		Converts a mixed unit (ft-in) string to a decimal number of inches.					
C_Mmatch		Returns the matching row number of a table, based on multiple criteria.					
C_RevStr		Reverses the order of a specified string.					
C_SheetName		Returns the worksheet name of the active cell.					
C_WorkbookName	Returns the workbook name of the active cell.					
