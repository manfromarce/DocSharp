using System.Collections.Generic;
using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(4002)]
    public class MasterTextPropAtom : Record
    {
        public List<MasterTextPropRun> MasterTextPropRuns = new List<MasterTextPropRun>();

        public MasterTextPropAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            while (this.Reader.BaseStream.Position < this.Reader.BaseStream.Length)
            {
                MasterTextPropRun m;
                m.count = this.Reader.ReadUInt32();
                m.indentLevel = this.Reader.ReadUInt16();
                this.MasterTextPropRuns.Add(m);
            }
        }
       
    }

    
    public struct MasterTextPropRun
    {
        public uint count;
        public ushort indentLevel;
    }

}
