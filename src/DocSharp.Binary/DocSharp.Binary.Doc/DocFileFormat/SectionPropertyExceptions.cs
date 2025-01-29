using DocSharp.Binary.CommonTranslatorLib;

namespace DocSharp.Binary.DocFileFormat
{
    public class SectionPropertyExceptions : PropertyExceptions
    {
        /// <summary>
        /// Parses the bytes to retrieve a SectionPropertyExceptions
        /// </summary>
        /// <param name="bytes">The bytes starting with the grpprl</param>
        public SectionPropertyExceptions(byte[] bytes)
            : base(bytes)
        {
        }

        #region IVisitable Members

        public override void Convert<T>(T mapping)
        {
            (mapping as IMapping<SectionPropertyExceptions>)?.Apply(this);
        }

        #endregion
    }
}
