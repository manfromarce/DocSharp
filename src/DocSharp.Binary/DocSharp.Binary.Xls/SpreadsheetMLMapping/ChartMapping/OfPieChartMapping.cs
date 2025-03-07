using DocSharp.Binary.OpenXmlLib.DrawingML;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using System;

namespace DocSharp.Binary.SpreadsheetMLMapping
{
    public class OfPieChartMapping : AbstractChartGroupMapping
    {
        public OfPieChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is BopPop))
            {
                throw new Exception("Invalid chart type");
            }

            var bopPop = crtSequence.ChartType as BopPop;

            // c:ofPieChart
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElOfPieChart, Dml.Chart.Ns);
            {

                // Axis Ids
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    writeValueElement(Dml.Chart.ElAxId, axisId.ToString());
                }
            }
            this._writer.WriteEndElement();
        }
        #endregion
    }
}
