﻿using System.Collections.Generic;
using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class PageSetupSequence : BiffRecordSequence, IVisitable
    {
        public Header Header;

        public Footer Footer;

        public HCenter HCenter;

        public VCenter VCenter;

        public LeftMargin LeftMargin;

        public RightMargin RightMargin;

        public TopMargin TopMargin;

        public BottomMargin BottomMargin;

        public Pls Pls;

        public List<Continue> Continues;

        public Setup Setup;

        public PageSetupSequence(IStreamReader reader)
            : base(reader)
        {
            //PAGESETUP = Header Footer HCenter VCenter [LeftMargin] [RightMargin] [TopMargin] [BottomMargin] [Pls *Continue] Setup

            // Header 
            this.Header = (Header)BiffRecord.ReadRecord(reader);

            // Footer 
            this.Footer = (Footer)BiffRecord.ReadRecord(reader);
            
            // HCenter 
            this.HCenter = (HCenter)BiffRecord.ReadRecord(reader); 
            
            // VCenter 
            this.VCenter = (VCenter)BiffRecord.ReadRecord(reader);
            
            // [LeftMargin] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.LeftMargin)
            {
                this.LeftMargin = (LeftMargin)BiffRecord.ReadRecord(reader);
            }
            
            // [RightMargin] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.RightMargin)
            {
                this.RightMargin = (RightMargin)BiffRecord.ReadRecord(reader);
            }
            
            // [TopMargin] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.TopMargin)
            {
                this.TopMargin = (TopMargin)BiffRecord.ReadRecord(reader);
            }
            
            // [BottomMargin] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.BottomMargin)
            {
                this.BottomMargin = (BottomMargin)BiffRecord.ReadRecord(reader);
            }
            
            // [Pls *Continue] 
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Pls)
            {
                this.Pls = (Pls)BiffRecord.ReadRecord(reader);

                this.Continues = new List<Continue>();
                while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
                {
                    this.Continues.Add((Continue)BiffRecord.ReadRecord(reader));
                }
            }
            
            // Setup
            this.Setup = (Setup)BiffRecord.ReadRecord(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            (mapping as IMapping<PageSetupSequence>)?.Apply(this);
        }

        #endregion
    }
}
