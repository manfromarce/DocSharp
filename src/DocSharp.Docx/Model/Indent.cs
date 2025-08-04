using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx.Model;

public class Indent
{
    public int? Left { get; set; }
    public int? LeftChars { get; set; }
    public int? Right { get; set; }
    public int? RightChars { get; set; }
    public int? Start { get; set; }
    public int? StartChars { get; set; }
    public int? End { get; set; }
    public int? EndChars { get; set; }
    public int? FirstLine { get; set; }
    public int? FirstLineChars { get; set; }
    public int? Hanging { get; set; }
    public int? HangingChars { get; set; }

    internal void SetFromAttribute(string targetAttribute, int val)
    {
        switch (targetAttribute.ToLowerInvariant())
        {
            // Only set values if they are still null, so that the correct priority is followed
            // (paragraph properties -> numbering properties -> style ...)
            case "left":
                Left ??= val;
                break;
            case "leftchars":
                LeftChars ??= val;
                break;
            case "right":
                Right ??= val;
                break;
            case "rightchars":
                RightChars ??= val;
                break;
            case "start":
                Start ??= val;
                break;
            case "startchars":
            case "startcharacters":
                StartChars ??= val;
                break;
            case "end":
                End ??= val;
                break;
            case "endchars":
            case "endcharacters":
                EndChars ??= val;
                break;
            case "firstline":
                FirstLine ??= val;
                break;
            case "firstlinechars":
                FirstLineChars ??= val;
                break;
            case "hanging":
                Hanging ??= val;
                break;
            case "hangingchars":
                HangingChars ??= val;
                break;
        }
    }
}
