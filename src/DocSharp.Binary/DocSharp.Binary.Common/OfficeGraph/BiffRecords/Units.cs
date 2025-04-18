using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.OfficeGraph
{
    /// <summary>
    /// This record MUST be zero, and MUST be ignored.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Units)]
    public class Units : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Units;

        public Units(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            reader.ReadBytes(2); // reserved

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
