using ExcelDna.Integration;
//using System.Globalization;

// ==================== START of NAMESPACE ====================
// note that this namespace is split across multiple CS files
//
// https://excel-dna.net/docs/introduction
//
namespace AKHudfC
{

    public class FuncInfo : XlCall
    // ==================== START of Class ====================
    // "XlCall" allows direct calling of Excel's native functions (I think).
    // public static class FuncInfo   
    // static classes are non-instantiable. Static classes cannot be inherited from another class.
    // --------------------------------------------------------------------------------
    {
        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tip
        [ExcelFunction(Description = "Straightforward method to get active worsheet name.")]
        public static string C_SheetName()
        {
			try
			{
				string strFullName = (string)XlCall.Excel(
					XlCall.xlfCell,
					"filename"
				);

				if (string.IsNullOrEmpty(strFullName))
					return "";

				int iBracket = strFullName.LastIndexOf(']');

				if (iBracket < 0 || iBracket >= strFullName.Length - 1)
					return "";

				return strFullName.Substring(iBracket + 1);
			}
			catch
			{
				return "";
			}		
        } // -------------------- End of Function --------------------

        // ==================== START of Function ====================
        // Description for IntelliSense Tool Tip
        [ExcelFunction(Description = "Straightforward method to get active workbook name.")]
        public static string C_WorkbookName()
        {

			try
			{
				string strFullName = (string)XlCall.Excel(
					XlCall.xlfCell,
					"filename"
				);

				if (string.IsNullOrEmpty(strFullName))
					return "";

				int iOpenBracket = strFullName.LastIndexOf('[');
				int iCloseBracket = strFullName.LastIndexOf(']');

				if (iOpenBracket < 0 || iCloseBracket < 0 || iCloseBracket <= iOpenBracket)
					return "";

				return strFullName.Substring(iOpenBracket + 1, iCloseBracket - iOpenBracket - 1);
			}
			catch
			{
				return "";
			}
			
        } // -------------------- End of Function --------------------


    } // ==================== END of Class ====================
} // ==================== END of Namespace ====================
