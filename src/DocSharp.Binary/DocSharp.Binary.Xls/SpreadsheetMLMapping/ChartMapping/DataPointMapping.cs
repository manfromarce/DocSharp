﻿using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.OpenXmlLib.DrawingML;

namespace DocSharp.Binary.SpreadsheetMLMapping
{
    public class DataPointMapping : AbstractChartMapping,
          IMapping<SsSequence>
    {
        int _index;

        public DataPointMapping(ExcelContext workbookContext, ChartContext chartContext, int index)
            : base(workbookContext, chartContext)
        {
            this._index = index;
        }

        public void Apply(SsSequence ssSequence)
        {
            // c:dPt
            this._writer.WriteStartElement(Dml.Chart.Prefix, "dPt", Dml.Chart.Ns);
            {
                // c:bubble3D
                // TODO

                // c:explosion
                // TODO

                // c:idx
                writeValueElement("idx", this._index.ToString());

                // c:invertIfNegative
                // TODO

                // c:marker
                // TODO

                // c:pictureOptions
                // TODO

                // c:spPr
                ssSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
            }
            this._writer.WriteEndElement();
        }
    }
}
