using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx;

public enum ShadingType
{
    Character,
    Paragraph,
    TableCell,
    TableRow
}

public enum BordersType
{
    Character,
    Paragraph,
    TableCell
}

[Flags]
public enum ConditionalFormattingFlags
{
    None = 0,
    FirstRow = 1 << 0,
    LastRow = 1 << 1,
    FirstColumn = 1 << 2,
    LastColumn = 1 << 3,
    OddRowBanding = 1 << 4,
    EvenRowBanding = 1 << 5,
    OddColumnBanding = 1 << 6,
    EvenColumnBanding = 1 << 7,
    NorthWestCell = 1 << 8, 
    NorthEastCell = 1 << 9, 
    SouthWestCell = 1 << 10, 
    SouthEastCell = 1 << 11, 
}