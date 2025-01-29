using DocSharp.Binary.StructuredStorage.Common;

namespace DocSharp.Binary.StructuredStorage.Writer
{
    /// <summary>
    /// Empty directory entry used to pad out directory stream.
    /// Author: math
    /// </summary>
    class EmptyDirectoryEntry : BaseDirectoryEntry
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="context">the current context</param>
        internal EmptyDirectoryEntry(StructuredStorageContext context)
            : base("", context)
        {
            this.Color = DirectoryEntryColor.DE_RED; // 0x0
            this.Type = DirectoryEntryType.STGTY_INVALID;            
        }

    }
}
