namespace DocSharp.Primitives;

public class PageMargins
{
	public double LeftMm { get; set; }
	public double TopMm { get; set; }
	public double RightMm { get; set; }
	public double BottomMm { get; set; }

	public PageMargins(double leftMm, double topMm, double rightMm, double bottomMm)
	{
		LeftMm = leftMm;
		TopMm = topMm;
		RightMm = rightMm;
		BottomMm = bottomMm;
	}

	public static PageMargins FromMillimeters(double leftMm, double topMm, double rightMm, double bottomMm)
		=> new PageMargins(leftMm, topMm, rightMm, bottomMm);

	public static PageMargins Uniform(double marginMm) => new PageMargins(marginMm, marginMm, marginMm, marginMm);

	public static PageMargins FromInches(double leftInch, double topInch, double rightInch, double bottomInch)
		=> FromMillimeters(leftInch * 25.4, topInch * 25.4, rightInch * 25.4, bottomInch * 25.4);

	public static PageMargins FromPoints(double leftPoints, double topPoints, double rightPoints, double bottomPoints)
		=> FromInches(leftPoints / 72.0, topPoints / 72.0, rightPoints / 72.0, bottomPoints / 72.0);
	
    // Default, Narrow and Large margins as defined in Microsoft Word.
    public static PageMargins Default => new PageMargins(20, 25, 20, 20);
	public static PageMargins Narrow => Uniform(12.7);
	public static PageMargins Large => new PageMargins(50.8, 25.4, 50.8, 25.4);

	public long LeftTwips() => UnitMetricHelper.ConvertToTwips(LeftMm, UnitMetric.Millimeter);
	public long TopTwips() => UnitMetricHelper.ConvertToTwips(TopMm, UnitMetric.Millimeter);
	public long RightTwips() => UnitMetricHelper.ConvertToTwips(RightMm, UnitMetric.Millimeter);
	public long BottomTwips() => UnitMetricHelper.ConvertToTwips(BottomMm, UnitMetric.Millimeter);
}
