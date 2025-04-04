﻿using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class AiSequence : BiffRecordSequence
    {
        public BRAI BRAI;

        public SeriesText SeriesText;

        public AiSequence(IStreamReader reader)
            : base(reader)
        {
            //AI = BRAI [SeriesText]

            //BRAI
            this.BRAI = (BRAI)BiffRecord.ReadRecord(reader);

            //[SeriesText]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.SeriesText)
            {
                this.SeriesText = (SeriesText)BiffRecord.ReadRecord(reader);
            }

        }
    }
}
