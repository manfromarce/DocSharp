using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class BiffRecordSequence
    {
        IStreamReader _reader;
        public IStreamReader Reader
        {
            get { return this._reader; }
            set { this._reader = value; }
        }

        public BiffRecordSequence(IStreamReader reader)
        {
            this._reader = reader;
        }
    }
}
