using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies a future record.
    /// </summary>
    public class FrtHeaderOld
    {
        /// <summary>
        /// An unsigned integer that specifies the record type identifier. 
        /// 
        /// MUST be identical to the record type identifier of the containing record.
        /// </summary>
        public RecordType rt;

        /// <summary>
        /// A FrtFlags that specifies attributes for this record. 
        /// 
        /// The value of grbitFrt.fFrtRef MUST be zero.
        /// </summary>
        public ushort grbitFrt;

        public FrtHeaderOld(IStreamReader reader)
        {
            this.rt = (RecordType)reader.ReadUInt16();
            this.grbitFrt = reader.ReadUInt16();
        }
    }
}
