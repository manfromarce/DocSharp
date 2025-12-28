using System.Globalization;

namespace DocSharp;

public class PageSize
{
    private UnitMetric unit = UnitMetric.Millimeter;

	public double WidthMm { get; set; }
	public double HeightMm { get; set; }

	public PageSize(double widthMm, double heightMm)
	{
		WidthMm = widthMm;
		HeightMm = heightMm;
	}

	public static PageSize FromMillimeters(double widthMm, double heightMm) => new PageSize(widthMm, heightMm);

	public static PageSize FromInches(double widthInch, double heightInch)
		=> new PageSize(widthInch * 25.4, heightInch * 25.4);

	public static PageSize FromPoints(double widthPoints, double heightPoints)
		=> FromInches(widthPoints / 72.0, heightPoints / 72.0);

	public static PageSize Default => RegionInfo.CurrentRegion.IsMetric ? A4 : Letter;

	// Standard sizes (portrait)
	public static PageSize A3 => new PageSize(297, 420);
	public static PageSize A4 => new PageSize(210, 297);
	public static PageSize A5 => new PageSize(148, 210);
	public static PageSize Letter => PageSize.FromInches(8.5, 11);
	public static PageSize Legal => PageSize.FromInches(8.5, 14);

    // Standard sizes (landscape)
	public static PageSize A3_Landscape => new PageSize(420, 297);
	public static PageSize A4_Landscape => new PageSize(297, 210);
	public static PageSize A5_Landscape => new PageSize(210, 148);
	public static PageSize Letter_Landscape => PageSize.FromInches(11, 8.5);
	public static PageSize Legal_Landscape => PageSize.FromInches(14, 8.5);

	public long WidthTwips() => UnitMetricHelper.ConvertToTwips(WidthMm, UnitMetric.Millimeter);
	public long HeightTwips() => UnitMetricHelper.ConvertToTwips(HeightMm, UnitMetric.Millimeter);

    public void SwapDimensions()
    {
        double temp = WidthMm;
        WidthMm = HeightMm;
        HeightMm = temp;
    }

    public void SetPortrait()
    {
        if (WidthMm > HeightMm)
            SwapDimensions();
    }

    public void SetLandscape()
    {
        if (HeightMm > WidthMm)
            SwapDimensions();
    }
}
