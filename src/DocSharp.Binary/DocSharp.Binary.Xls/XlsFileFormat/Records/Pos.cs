﻿using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the size and position for a legend, an attached label, 
    /// or the plot area, as specified by the primary axis group. <br/>
    /// This record MUST be ignored for the plot area when the fManPlotArea field of ShtProps is set to 1.
    /// </summary>
    [BiffRecord(RecordType.Pos)]
    public class Pos : BiffRecord
    {
        public enum PositionMode
        {
            MDFX = 0,
            MDABS = 1,
            MDPARENT = 2,
            MDKTH = 3,
            MDCHART = 5
        }

        public const RecordType ID = RecordType.Pos;

        /// <summary>
        /// A PositionMode that specifies the positioning mode for the upper-left corner of a legend, 
        /// an attached label, or the plot area. 
        /// </summary>
        public PositionMode mdTopLt;

        /// <summary>
        /// A PositionMode that specifies the positioning mode for the lower-right corner of a legend, 
        /// an attached label, or the plot area.
        /// </summary>
        public PositionMode mdBotRt;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public short x1;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public short y1;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public short x2;

        /// <summary>
        /// A signed integer that specifies a position. <br/>
        /// The meaning is specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
        /// </summary>
        public short y2;

        public Pos(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.mdTopLt = (PositionMode)reader.ReadInt16();
            this.mdBotRt = (PositionMode)reader.ReadInt16();
            this.x1 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.y1 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.x2 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.y2 = reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
