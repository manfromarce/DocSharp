using System.Text;
using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.DataContainer;

using DocSharp.Binary.StructuredStorage.Reader; 

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class XlsDocument :  IVisitable
    {

#if !NETFRAMEWORK
        static XlsDocument()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
#endif

        /// <summary>
        /// Some constant strings 
        /// </summary>
        private const string WORKBOOK = "Workbook";
        private const string ALTERNATE1 = "Book"; 

        /// <summary>
        /// The workbook streamreader 
        /// </summary>
        private VirtualStreamReader workBookStreamReader; 

        /// <summary>
        /// The Workbookextractor / container 
        /// </summary>
        private WorkbookExtractor workBookExtr;

        /// <summary>
        /// This attribute stores the hole Workbookdata 
        /// </summary>
        public WorkBookData WorkBookData;

        /// <summary>
        /// The StructuredStorageFile itself
        /// </summary>
        public StructuredStorageReader Storage;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="file"></param>
        public XlsDocument(StructuredStorageReader reader)
        {
            this.WorkBookData = new WorkBookData();
            this.Storage = reader;

            if (reader.FullNameOfAllStreamEntries.Contains("\\" + WORKBOOK))
            {
                this.workBookStreamReader = new VirtualStreamReader(reader.GetStream(WORKBOOK));
            }
            else if (reader.FullNameOfAllStreamEntries.Contains("\\" + ALTERNATE1))
            {
                this.workBookStreamReader = new VirtualStreamReader(reader.GetStream(ALTERNATE1));
            }
            else
            {
                throw new ExtractorException(ExtractorException.WORKBOOKSTREAMNOTFOUND);
            }

            this.workBookExtr = new WorkbookExtractor(this.workBookStreamReader, this.WorkBookData); 
        }


        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            (mapping as IMapping<XlsDocument>)?.Apply(this);
        }

        #endregion
    }
}
