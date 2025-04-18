﻿using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies properties of a chart group and specifies the beginning 
    /// of a collection of records as defined by the Chart Sheet Substream ABNF. 
    /// The collection of records specifies a chart group.
    /// </summary>
    [BiffRecord(RecordType.ChartFormat)]
    public class ChartFormat : BiffRecord
    {
        public const RecordType ID = RecordType.ChartFormat;

        /// <summary>
        /// A bit that specifies whether the color for each data point and the color 
        /// and type for each data marker varies. If the chart group has multiple series, 
        /// or the chart group has one series and the type is a surface, stock, or area 
        /// chart group, then this field MUST be ignored, and the data points do not vary. 
        /// For all other chart group types, if the chart group has one series, then a value 
        /// of 0x1 specifies that the data points vary. 
        /// 
        /// MUST be a value from the following table: 
        ///     Value         Meaning
        ///     0x0           The color for each data point and the color and type for each data marker does not vary.
        ///     0x1           The color for data points or the color or type for data markers varies.
        /// </summary>
        public bool fVaried;

        /// <summary>
        /// An unsigned integer that specifies the drawing order of the chart group relative 
        /// to the other chart groups, where 0x0000 is the bottom of the z-order. 
        /// 
        /// This value MUST be unique for each instance of this record and MUST be less than or equal to 0x0009.
        /// </summary>
        public ushort icrt;

        public int[] AxisIds;

        /// <summary>
        /// An unsigned integer that specifies the zero-based index of this ChartFormat 
        /// record in the collection of ChartFormat records in the current chart sheet substream.
        /// 
        /// NOTE: This information is added at parse time and is not stored in the binary file format.
        /// </summary>
        public ushort idx;

        public ChartFormat(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            
            //ignore beginning of record
            reader.ReadBytes(16);
            this.fVaried = Utils.BitmaskToBool(reader.ReadUInt16(), 0x0001);
            this.icrt = reader.ReadUInt16();

            this.AxisIds = ChartAxisIdGenerator.Instance.AxisIds;
            this.idx = ChartFormatIdGenerator.Instance.GenerateId();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
