using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer.Models.Tables.Grids
{
    internal class GridRow
    {
        public GridRow(double height, HeightRuleValues heightRule)
        {
            this.Height = height;
            this.HeightRule = heightRule;
        }

        public double Height { get; private set; }

        public HeightRuleValues HeightRule { get; }

        public void Expand(double height)
        {
            this.Height = Math.Max(height, this.Height);
        }
    }
}
