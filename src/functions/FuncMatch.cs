using ExcelDna.Integration;
//using static ExcelDna.Integration.XlCall; // You don’t have to have your class inherit from XlCall. That can be better done these days by adding <that>.
//using System.Reflection; // required for .GetValue Method
using System; // this is needed for TYPE definitions
using System.Collections.Generic; // need this for LIST operations
using System.Text.RegularExpressions;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{

    public class FuncMatch : XlCall
    // ==================== START of Class ====================
    // public static class FuncMatch 
    // --------------------------------------------------------------------------------
    {
        // ==================== START of Function ====================
        // A MATCH function for multiple criteria
        // This is indended to / could be returned from Excel's native MATCH function
        // make these arguments DOUBLE, not INT
        //============================================================
        [ExcelFunction(IsMacroType = false, IsVolatile = false,
        Description = "Returns row no. of first (exact) MATCH of multiple (3) COLUMN criteria.")]
        public static object C_MMatch
            ([ExcelArgument(AllowReference = true, Name="MatchRange1", Description ="1st range to search.")]
            object[] rngMatch1,
            [ExcelArgument(AllowReference = false, Name="Criteria1", Description ="Criteria for 1st range.")]
            object objCrit1,
            //
            // CONSIDER THESE TO BE OPTIONAL ARGUMENTS - ONLY ONE REQUIRED RANGE/CRITERIA FOR EASIER COMPARISON AGAINS NATIVE MATCH FUNCTION
            //			
            [ExcelArgument(AllowReference = true, Name="[MatchRange2]", Description ="Optional 2nd range to search. Length must equal other ranges.")]
            object[] rngMatch2,
            [ExcelArgument(AllowReference = false, Name="[Criteria2]", Description ="Optional criteria for 2nd range.")]
            object objCrit2,
            [ExcelArgument(AllowReference = true, Name="[MatchRange3]", Description ="Optional 3rd range to search. Length must equal other ranges.")]
            object[] rngMatch3,
            [ExcelArgument(AllowReference = false, Name="[Criteria3]", Description ="Optional criteria for 3rd range.")]
            object objCrit3,
            [ExcelArgument(AllowReference = true, Name="[MatchRange4", Description ="Optional 4th rangeto search. Length must equal other ranges.")]
            object[] rngMatch4,
            [ExcelArgument(AllowReference = false, Name="[Criteria4]", Description ="Optional criteria for 4th range.")]
            object objCrit4
            )
        {
            // ------------------------------------------------------------
            // Preliminary work to filter out unused, optional arguments
            // ------------------------------------------------------------

            // Check for OPTIONAL ARGUMENTS, set to NULL if not present
            rngMatch2 = ToolsArgs.CheckOpt(rngMatch2, "NULL");
			rngMatch3 = ToolsArgs.CheckOpt(rngMatch3, "NULL");
            rngMatch4 = ToolsArgs.CheckOpt(rngMatch4, "NULL");
			
			objCrit2 = ToolsArgs.CheckOpt(objCrit2, "NULL");
            objCrit3 = ToolsArgs.CheckOpt(objCrit3, "NULL");
            objCrit4 = ToolsArgs.CheckOpt(objCrit4, "NULL");

            // Create a List of criteria  
            // Create a List of ranges/arrays
            List<object> objListCrit = new List<object> { objCrit1, objCrit2, objCrit3, objCrit4 };
            List<object[]> objListRng = new List<object[]> { rngMatch1, rngMatch2, rngMatch3, rngMatch4 };
			
			// Remove omitted optional pairs, working backward
			for (int i = objListCrit.Count - 1; i >= 1; i--)
			{
				if (ToolsArgs.GetStr(objListCrit[i]) == "NULL")
				{
					objListCrit.RemoveAt(i);
					objListRng.RemoveAt(i);
				}
			}

			// Check range lengths after omitted arguments are removed
			int iRowCount = objListRng[0].GetLength(0);

			for (int i = 1; i < objListRng.Count; i++)
			{
				if (objListRng[i].GetLength(0) != iRowCount)
				{
					return ToolsErrors.GetErrorNA();
				}
			}

            //combine 1d range arrays into a single 2d array
            //builds a WIDE array...may need to be inverted later
            //int iRows = rngMatch1.GetLength(0);
            int iRows = objListRng[0].GetLength(0);
            int iCols = objListCrit.Count;
            object[,] arrRng = new object[iRows, iCols];
            for (int i = 0; i < iCols; i++)
            {
                for (int j = 0; j < iRows; j++)
                {
                    arrRng[j, i] = objListRng[i][j];
                }

            }

            // i = 0; j = 0; k = 0;
            // int iLimit = arrRng.GetLength(0) - 1; // length of each loop (number of rows/columns)
            // int jLimit = arrRng.GetLength(1) - 1; // number of criteria to loop through
			
			//int iRows = objListRng[0].GetLength(0);
			//int iCols = objListCrit.Count;

			//string strRngItem = null;
			//string strCriteria = null;
			//double? dblRngItem = null;
			//double? dblCriteria = null;
			//bool bMatch = false;
			
			// STOPPED UPDATING HERE AT 10AM 22 MAY 2026

			// ------------------------------------------------------------
			// Start Comparisons
			// Loop through each row/item in the match ranges.
			// For each row, check all active criteria.
			// ------------------------------------------------------------

			for (int i = 0; i < iRows; i++)
			{
				bool bAllCriteriaMatch = true;

				for (int j = 0; j < iCols; j++)
				{
					object objRngItem = objListRng[j][i];
					object objCriteria = objListCrit[j];

					var varRngItem = ToolsArgs.ObjGet(objRngItem);
					var varCriteria = ToolsArgs.ObjGet(objCriteria);

					bool bThisCriteriaMatch = false;

					// ------------------------------------------------------------
					// Do individual comparison checks here.
					// Set bThisCriteriaMatch = true when this criteria matches.
					// ------------------------------------------------------------
					if (ToolsArgs.ObjMatch(objRngItem, objCriteria))
					{
						bThisCriteriaMatch = true;
					}

					// If any criteria fails, this row is not a match.
					if (!bThisCriteriaMatch)
					{
						bAllCriteriaMatch = false;
						break;
					}
				}

				// If every criteria matched, return Excel-style 1-based position.
				if (bAllCriteriaMatch)
				{
					return i + 1;
				}
			}

			// No match found.
			return ToolsErrors.GetErrorNA();


        } // -------------------- End of Function --------------------

    } // ==================== END of Class ====================
} // ==================== END of Namespace ====================
