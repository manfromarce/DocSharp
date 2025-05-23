using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;

namespace DocSharp.Binary.OfficeGraph
{
    /// <summary>
    /// This record specifies that the chart group is a line chart group and specifies the chart group attributes.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Line)]
    public class Line : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Line;

        /// <summary>
        /// A bit that specifies whether the data points in the chart group that share the same 
        /// category (3) are stacked one on top of the next.
        /// </summary>
        public bool fStacked;

        /// <summary>
        /// A bit that specifies whether the data points in the chart group are displayed as a percentage 
        /// of the sum of all data points in the chart group that share the same category (3).<br/>
        /// MUST be false if fStacked is false.
        /// </summary>
        public bool f100;

        /// <summary>
        /// A bit that specifies whether one or more data markers in the chart group has shadows.
        /// </summary>
        public bool fHasShadow;

        public Line(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fStacked = Utils.BitmaskToBool(flags, 0x1);
            this.f100 = Utils.BitmaskToBool(flags, 0x2);
            this.fHasShadow = Utils.BitmaskToBool(flags, 0x4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
