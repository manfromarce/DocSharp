using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx;

internal class PictureProperties
{
    internal long Width { get; set; }
    internal long Height { get; set; }
    internal long WidthGoal { get; set; }
    internal long HeightGoal { get; set; }

    internal long CropLeft { get; set; } = 0;
    internal long CropRight { get; set; } = 0;
    internal long CropTop { get; set; } = 0;
    internal long CropBottom { get; set; } = 0;
}
