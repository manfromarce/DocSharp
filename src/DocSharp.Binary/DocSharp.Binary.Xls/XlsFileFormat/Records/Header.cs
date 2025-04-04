using System.Diagnostics;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the header text of the current sheet when printed.
    /// </summary>
    [BiffRecord(RecordType.Header)] 
    public class Header : BiffRecord
    {
        public const RecordType ID = RecordType.Header;

        /// <summary>
        /// An XLUnicodeString that specifies the header text for the current sheet. 
        /// It is optional and exists only if the record size is not zero. The text 
        /// appears at the top of every page when printed. The length of the text MUST 
        /// be less than or equal to 255. The header text can contain special commands, 
        /// for example a placeholder for the page number, current date or text formatting attributes.
        /// </summary>
        public XLUnicodeString headerText; 

        public Header(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            if (this.Length > 0)
            {
                this.headerText = new XLUnicodeString(reader);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
