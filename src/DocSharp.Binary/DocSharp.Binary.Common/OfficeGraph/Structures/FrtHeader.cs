using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.OfficeGraph
{
    /// <summary>
    /// This structure specifies a future record.
    /// </summary>
    public class FrtHeader
    {
        /// <summary>
        /// An unsigned integer that specifies the record type identifier. 
        /// 
        /// MUST be identical to the record type identifier of the containing record.
        /// </summary>
        public ushort rt;

        /// <summary>
        /// A FrtFlags that specifies attributes for this record. 
        /// 
        /// The value of grbitFrt.fFrtRef MUST be zero.
        /// </summary>
        public ushort grbitFrt;

        public FrtHeader(IStreamReader reader)
        {
            this.rt = reader.ReadUInt16();
            this.grbitFrt = reader.ReadUInt16();

            // ignore remaing record data
            reader.ReadBytes(8);
        }
    }
}
