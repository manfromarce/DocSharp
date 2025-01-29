using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class ChartSheetSequence : BiffRecordSequence, IVisitable
    {
        public BOF BOF;

        public ChartFrtInfo ChartFrtInfo;

        public ChartSheetContentSequence ChartSheetContentSequence;

        public ChartSheetSequence(IStreamReader reader) : base(reader)
        {
            //BOF 
            this.BOF = (BOF)BiffRecord.ReadRecord(reader);

            // [ChartFrtInfo] (not specified)
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ChartFrtInfo)
            {
                this.ChartFrtInfo = (ChartFrtInfo)BiffRecord.ReadRecord(reader);
            }

            //CHARTSHEETCONTENT
            this.ChartSheetContentSequence = new ChartSheetContentSequence(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            (mapping as IMapping<ChartSheetSequence>)?.Apply(this);
        }

        #endregion
    }
}