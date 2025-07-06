namespace DocSharp.Renderer.Models.Sections.Columns
{
    internal class ColumnConfig
    {
        public ColumnConfig(double width, double space)
        {
            this.Width = width;
            this.Space = space;
        }

        public double Width { get; }
        public double Space { get; }
    }
}
