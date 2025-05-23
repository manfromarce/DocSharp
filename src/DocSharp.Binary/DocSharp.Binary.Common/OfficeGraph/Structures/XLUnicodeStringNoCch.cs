using System.Text;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;

namespace DocSharp.Binary.OfficeGraph
{
    /// <summary>
    /// This structure specifies a Unicode string. 
    /// 
    /// When an XLUnicodeStringNoCch is used, the count of characters in the string 
    /// MUST be specified in the structure that uses the XLUnicodeStringNoCch.
    /// </summary>
    public class XLUnicodeStringNoCch
    {
        /// <summary>
        /// A bit that specifies whether the characters in rgb are double-byte characters. 
        /// 
        /// MUST be a value from the following table: 
        ///     Value   Meaning
        ///     0x0     All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
        ///     0x1     All the characters in the string are saved as double-byte characters in rgb.
        /// </summary>
        public bool fHighByte;

        /// <summary>
        /// An array of bytes that specifies the characters. 
        /// 
        /// If fHighByte is 0x0, the size of the array MUST be equal to the count of characters in the string. 
        /// If fHighByte is 0x1, the size of the array MUST be equal to 2 times the count of characters in the string.
        /// </summary>
        public byte[] rgb;

        
        public XLUnicodeStringNoCch(IStreamReader reader, ushort cch)
        {
            this.fHighByte = Utils.BitmaskToBool(reader.ReadByte(), 0x0001);

            if (this.fHighByte)
            {
                this.rgb = new byte[2 * cch];
            }
            else
            {
                this.rgb = new byte[cch];
            }

            this.rgb = reader.ReadBytes(this.rgb.Length);
        }

        public string Value
        {
            get
            {
                if (this.fHighByte)
                {
                    return Encoding.Unicode.GetString(this.rgb);
                }
                else
                {
                    return Encoding.GetEncoding(1252).GetString(this.rgb);
                }
            }
        }
    }
}
