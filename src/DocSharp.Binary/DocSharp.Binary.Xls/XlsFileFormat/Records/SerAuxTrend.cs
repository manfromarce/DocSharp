﻿using System.Diagnostics;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies a trendline.
    /// </summary>
    [BiffRecord(RecordType.SerAuxTrend)]
    public class SerAuxTrend : BiffRecord
    {
        public enum TrendlineType
        {
            Polynomial,
            Exponential,
            Logarithmic,
            Power,
            MovingAverage
        }

        public const RecordType ID = RecordType.SerAuxTrend;

        /// <summary>
        /// Specifies the type of trendline.
        /// </summary>
        public TrendlineType regt;

        /// <summary>
        /// specifies the polynomial order or moving average period. <br/>
        /// MUST be greater than or equal to 0x02 and less than or equal to 0x06 if regt equals 0x00; <br/>
        /// MUST be greater than or equal to 0x02 and less than or equal to the value of the cValx field of 
        /// the Series record specified by the preceding SerParent record minus one if regt equals 0x04. <br/>
        /// MUST be ignored for other values of regt.
        /// </summary>
        public byte ordUser;

        /// <summary>
        /// Specifies where the trendline intersects the value Axis or vertical Axis on bubble and scatter chart groups. <br/>
        /// If no intercept is specified, this ChartNumNillable MUST be null
        /// </summary>
        public double? numIntercept;

        /// <summary>
        /// A Boolean that specifies whether the trendline equation is displayed in the trendline label.<br/> 
        /// MUST be ignored if regt equals 0x04. <br/>
        /// MUST be ignored if the chart sheet substream contains an attached label with an 
        /// ObjectLink record that contains both a wLinkObj field equal to 0x0004 and 
        /// a wLinkVar1 field equal to the zero-based index into a Series record in the 
        /// collection of Series records in the current chart sheet substream that represents 
        /// this trendline, and the attached label contains a SeriesText record.
        /// </summary>
        public bool fEquation;

        /// <summary>
        /// A Boolean that specifies whether the R-squared value is displayed in the trendline label.<br/> 
        /// MUST be ignored if regt equals 0x04.<br/>
        /// MUST be ignored if the chart sheet substream contains an attached label with an ObjectLink 
        /// record that contains both a wLinkObj field equal to 0x0004 and a wLinkVar1 field equal to 
        /// the zero-based index into a Series record in the collection of Series records in the 
        /// current chart sheet substream that represents this trendline, and the attached label 
        /// contains a SeriesText record.
        /// </summary>
        public bool fRSquared;

        /// <summary>
        /// An Xnum that specifies the number of periods to forecast forward.
        /// </summary>
        public double numForecast;

        /// <summary>
        /// An Xnum that specifies the number of periods to forecast backward.
        /// </summary>
        public double numBackcast;

        public SerAuxTrend(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.regt = (TrendlineType)reader.ReadByte();
            this.ordUser = reader.ReadByte();

            //read the nullable double value (ChartNumNillable)
            this.numIntercept = new ChartNumNillable(reader).value;
            
            this.fEquation = Utils.ByteToBool(reader.ReadByte());
            this.fRSquared = Utils.ByteToBool(reader.ReadByte());
            this.numForecast = reader.ReadDouble();
            this.numBackcast = reader.ReadDouble();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
