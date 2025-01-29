namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.StyleData
{
    public class BorderData
    {
        public BorderPartData top;
        public BorderPartData bottom;
        public BorderPartData left;
        public BorderPartData right;
        public BorderPartData diagonal;

        public ushort diagonalValue; 

        public BorderData()
        {

            this.diagonalValue = 0; 
        }

        /// <summary>
        /// Equals Method 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj is BorderData bd)
            {
                if (this.top?.Equals(bd.top) == true && 
                    this.bottom?.Equals(bd.bottom) == true && 
                    this.left?.Equals(bd.left) == true &&
                    this.right?.Equals(bd.right) == true &&
                    this.diagonal?.Equals(bd.diagonal) == true && 
                    this.diagonalValue == bd.diagonalValue)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
