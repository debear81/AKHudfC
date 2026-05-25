using ExcelDna.Integration;

public static class InterpolationFunctions
{
	// May 2026 - altered user input order to be more logical
	[ExcelFunction(Description = "Linear interpolation: returns Y at X between points (X1,Y1) and (X2,Y2).")]
	public static object C_Linterp(
		[ExcelArgument(Description = "X value to interpolate at")] object x,
		[ExcelArgument(Description = "Point 1 X value")] object x1,
		[ExcelArgument(Description = "Point 1 Y value")] object y1,
		[ExcelArgument(Description = "Point 2 X value")] object x2,
		[ExcelArgument(Description = "Point 2 Y value")] object y2)
	{
		// Try to parse inputs as doubles
		if (!TryGetDouble(x, out double dx) ||
			!TryGetDouble(x1, out double dx1) ||
			!TryGetDouble(y1, out double dy1) ||
			!TryGetDouble(x2, out double dx2) ||
			!TryGetDouble(y2, out double dy2))
		{
			return ExcelError.ExcelErrorValue;
		}

		// Handle boundary conditions
		if (dx1 == dx2)
			return ExcelError.ExcelErrorDiv0;

		return dy1 + (dx - dx1) * (dy2 - dy1) / (dx2 - dx1);
	}
	
	

    // Helper method to safely convert Excel inputs to double
	// Slightly beefed-up, May 2026, to rule out divide by zero possibility
	private static bool TryGetDouble(object input, out double result)
	{
		switch (input)
		{
			case double d:
				result = d;
				return true;

			case int i:
				result = i;
				return true;

			case string s when double.TryParse(s, out double parsed):
				result = parsed;
				return true;

			default:
				result = double.NaN;
				return false;
		}
	}
	
	
}
