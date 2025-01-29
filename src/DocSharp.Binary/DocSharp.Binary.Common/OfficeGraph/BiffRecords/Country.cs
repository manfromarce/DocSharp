using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.OfficeGraph
{
    /// <summary>
    /// This record specifies the country value that is unused and MUST be ignored.
    /// </summary>
    [OfficeGraphBiffRecordAttribute(GraphRecordNumber.Country)]
    public class Country : OfficeGraphBiffRecord
    {
        public const GraphRecordNumber ID = GraphRecordNumber.Country;

        public Country(IStreamReader reader, GraphRecordNumber id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // content is completely ignored
            reader.ReadBytes(4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
