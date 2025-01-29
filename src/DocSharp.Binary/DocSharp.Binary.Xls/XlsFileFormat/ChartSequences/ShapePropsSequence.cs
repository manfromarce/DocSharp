using System.Collections.Generic;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class ShapePropsSequence : BiffRecordSequence
    {
        public ShapePropsStream ShapePropsStream;

        public List<ContinueFrt12> ContinueFrt12s;

        public ShapePropsSequence(IStreamReader reader)
            : base(reader)
        {
            // SHAPEPROPS = ShapePropsStream *ContinueFrt12

            // ShapePropsStream
            this.ShapePropsStream = (ShapePropsStream)BiffRecord.ReadRecord(reader);

            // *ContinueFrt12
            this.ContinueFrt12s = new List<ContinueFrt12>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ContinueFrt12)
            {
                this.ContinueFrt12s.Add((ContinueFrt12)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
