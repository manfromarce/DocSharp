using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.OpenXmlLib.DrawingML;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;

namespace DocSharp.Binary.SpreadsheetMLMapping
{
    public abstract class AbstractAxisMapping : AbstractChartMapping,
          IMapping<AxesSequence>
    {
        public AbstractAxisMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AxesSequence> Members

        public virtual void Apply(AxesSequence axesSequence)
        {
            // EG_AxShared
            writeValueElement(Dml.Chart.ElAxId, axesSequence.IvAxisSequence.Axis.AxisId.ToString());
        }

        #endregion
    }
}
