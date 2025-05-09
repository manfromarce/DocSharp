﻿using System.Collections.Generic;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class SortData12Sequence : BiffRecordSequence
    {
        public SortData SortData;

        public List<ContinueFrt12> ContinueFrt12s;

        public SortData12Sequence(IStreamReader reader)
            : base(reader)
        {
            // SORTDATA12 = SortData *ContinueFrt12

            // SortData
            this.SortData = (SortData)BiffRecord.ReadRecord(reader);

            // *ContinueFrt12
            this.ContinueFrt12s = new List<ContinueFrt12>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ContinueFrt12)
            {
                this.ContinueFrt12s.Add((ContinueFrt12)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
