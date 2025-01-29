using System.Collections.Generic;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;

namespace DocSharp.Binary.DocFileFormat
{
    public class AnnotationOwnerList : List<string>
    {
        public AnnotationOwnerList(FileInformationBlock fib, VirtualStream tableStream) : base()
        {
            tableStream.Seek(fib.fcGrpXstAtnOwners, System.IO.SeekOrigin.Begin);

            while (tableStream.Position < (fib.fcGrpXstAtnOwners + fib.lcbGrpXstAtnOwners))
            {
                this.Add(Utils.ReadXst(tableStream));
            }
        }
    }
}
