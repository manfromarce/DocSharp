﻿using System.Diagnostics;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies rich text formatting within chart titles, trendline, and data labels.
    /// </summary>
    [BiffRecord(RecordType.AlRuns)]
    public class AlRuns : BiffRecord
    {
        public const RecordType ID = RecordType.AlRuns;

        /// <summary>
        /// An unsigned integer that specifies the number of rich text runs. 
        /// MUST be greater than or equal to 3 and less than or equal to 256.
        /// </summary>
        public ushort cRuns;

        public FormatRun[] rgRuns;

        public AlRuns(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.cRuns = reader.ReadUInt16();

            if (this.cRuns > 0)
            {
                this.rgRuns = new FormatRun[this.cRuns];

                for (int i = 0; i < this.cRuns; i++)
                {
                    this.rgRuns[i] = new FormatRun(reader);
                }
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
