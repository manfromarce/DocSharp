﻿using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgName : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgName;

        public int nameindex;

        public PtgName(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 5;
            this.Data = "";
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.nameindex = this.Reader.ReadInt32(); 
        }
    }
}

