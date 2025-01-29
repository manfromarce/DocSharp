using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class LdSequence : BiffRecordSequence, IVisitable
    {
        public Legend Legend;

        public Begin Begin;

        public Pos Pos;

        public AttachedLabelSequence AttachedLabelSequence;

        public FrameSequence FrameSequence;

        public CrtLayout12 CrtLayout12;

        public TextPropsSequence TextPropsSequence;

        public CrtMlfrtSequence CrtMlfrtSequence;

        public End End; 
        
        public LdSequence(IStreamReader reader)
            : base(reader)
        {

            /// Legend Begin Pos ATTACHEDLABEL [FRAME] [CrtLayout12] [TEXTPROPS] [CRTMLFRT] End

            // Legend 
            this.Legend = (Legend)BiffRecord.ReadRecord(reader); 

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // Pos 
            this.Pos = (Pos)BiffRecord.ReadRecord(reader);

            // [ATTACHEDLABEL]
            this.AttachedLabelSequence = new AttachedLabelSequence(reader); 

            // [FRAME]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.Frame)
            {
                this.FrameSequence = new FrameSequence(reader);
            }

            // [CrtLayout12]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtLayout12)
            {
                this.CrtLayout12 = (CrtLayout12)BiffRecord.ReadRecord(reader);
            }

            // [TEXTPROPS]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.RichTextStream
                || BiffRecord.GetNextRecordType(reader) == RecordType.TextPropsStream)
            {
                this.TextPropsSequence = new TextPropsSequence(reader);
            }

            //[CRTMLFRT]
            if (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlfrtSequence = new CrtMlfrtSequence(reader);
            }

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader); 

        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            (mapping as IMapping<LdSequence>)?.Apply(this);
        }

        #endregion
    }
}
