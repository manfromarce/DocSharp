using System;
using System.Diagnostics;
using System.Globalization;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgNum : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgNum;

        public PtgNum(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 9;
            this.Data = Convert.ToString(this.Reader.ReadDouble(), CultureInfo.GetCultureInfo("en-US")); 

            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
