using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocSharp.Rtf
{
    /// <summary>
    /// Represents a source of RTF content. 
    /// It can be created from Stream, file path, TextReader or content string.
    /// Using a <see cref="Stream"/> or file path (which is handled as FileStream) 
    /// is preferred as an RTF file can switch binary encodings in the middle of a file.
    /// Please note that a complete RTF document should be used in any case,
    /// as partial strings of RTF syntax are not guaranteed to be converted properly.
    /// </summary>
    public class RtfSource
    {
        static RtfSource()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public Document RtfDocument { get; set; }

        /// <summary>
        /// Create an RTF source from a text reader.
        /// </summary>
        /// <param name="reader">The <see cref="TextReader"/> to use</param>
        public RtfSource(TextReader reader)
        {
            var parser = new Parser(reader);
            RtfDocument = parser.Parse();
            // In this case the application is responsible for closing the reader.
        }

        /// <summary>
        /// Create an RTF source from a Stream
        /// </summary>
        /// <param name="stream">The <see cref="Stream"/> to use</param>
        public RtfSource(Stream stream)
        {
            using (var reader = new RtfStreamReader(stream))
            {
                var parser = new Parser(reader);
                RtfDocument = parser.Parse();
                // In this case the application is responsible for disposing the stream.
            }
        }

        /// <summary>
        /// Create an RTF source from a Stream
        /// </summary>
        /// <param name="stream">The <see cref="Stream"/> to use</param>
        public static RtfSource FromStream(Stream stream)
        {
            return new RtfSource(stream);
        }

        /// <summary>
        /// Create an RTF source from a file path
        /// </summary>
        /// <param name="filePath">The file path to load</param>
        public static RtfSource FromFile(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return new RtfSource(fs);
            }
        }

        /// <summary>
        /// Create an RTF source from a string containing RTF syntax.
        /// </summary>
        /// <param name="rtfContent">The RTF string to use</param>
        public static RtfSource FromRtfString(string rtfContent)
        {
            using (var sr = new StringReader(rtfContent))
            {
                return new RtfSource(sr);
            }
        }
    }
}
