using DocSharp.Binary.CommonTranslatorLib;
namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class ChartSheetData : SheetData
    {
        public ChartSheetSequence ChartSheetSequence;

        public ChartSheetData()
        {
        }

        public override void Convert<T>(T mapping)
        {
            (mapping as IMapping<ChartSheetSequence>)?.Apply(this.ChartSheetSequence);
        }
    }
}
