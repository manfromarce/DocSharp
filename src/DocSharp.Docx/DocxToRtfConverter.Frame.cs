using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessFrameProperties(FrameProperties fp, StringBuilder sb)
    {
        if (fp.Width?.Value != null && int.TryParse(fp.Width.Value, out int w))
        {
            sb.Append($"\\absw{w}");
        }
        if (fp.HeightType != null)
        {
            if (fp.HeightType.Value == HeightRuleValues.Auto)
            {
                sb.Append("\\absh0");
            }
            else if (fp.Height != null && fp.Height.HasValue)
            {
                if (fp.HeightType.Value == HeightRuleValues.AtLeast)
                {
                    sb.Append($"\\absh{fp.Height.Value}");
                }
                else
                {
                    sb.Append($"\\absh-{fp.Height.Value}");
                }
            }
        }
        if (fp.HorizontalPosition?.Value != null)
        {
            if (fp.HorizontalPosition.Value == HorizontalAnchorValues.Margin)
            {
                sb.Append(@"\phmrg");
            }
            else if (fp.HorizontalPosition.Value == HorizontalAnchorValues.Page)
            {
                sb.Append(@"\phpg");
            }
            else if (fp.HorizontalPosition.Value == HorizontalAnchorValues.Text)
            {
                sb.Append(@"\phcol");
            }
        }
        if (fp.XAlign?.Value != null)
        {
            if (fp.XAlign.Value == HorizontalAlignmentValues.Center)
            {
                sb.Append("\\posxc");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Inside)
            {
                sb.Append("\\posxi");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Outside)
            {
                sb.Append("\\posxo");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Left)
            {
                sb.Append("\\posxl");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Right)
            {
                sb.Append("\\posxr");
            }           
        }
        if (fp.X?.Value != null && int.TryParse(fp.X.Value, out int x))
        {
            if (x > 0)
                sb.Append($"\\posx{x}");
            else
                sb.Append($"\\posnegx{x}");
        }
        if (fp.HorizontalSpace?.Value != null && int.TryParse(fp.HorizontalSpace?.Value, out int h))
        {
            sb.Append($"\\dfrmtxtx{h}");
        }
        if (fp.VerticalPosition?.Value != null)
        {
            if (fp.VerticalPosition.Value == VerticalAnchorValues.Margin)
            {
                sb.Append(@"\pvmrg");
            }
            else if (fp.VerticalPosition.Value == VerticalAnchorValues.Page)
            {
                sb.Append(@"\pvpg");
            }
            else if (fp.VerticalPosition.Value == VerticalAnchorValues.Text)
            {
                sb.Append(@"\pvpara");
            }
        }
        if (fp.YAlign?.Value != null)
        {
            if (fp.YAlign.Value == VerticalAlignmentValues.Bottom)
            {
                sb.Append("\\posyb");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Center)
            {
                sb.Append("\\posyc");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Inline)
            {
                sb.Append("\\posyil");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Inside)
            {
                sb.Append("\\posyin");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Outside)
            {
                sb.Append("\\posyout");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Top)
            {
                sb.Append("\\posyt");
            }
        }
        if (fp.Y?.Value != null && int.TryParse(fp.Y.Value, out int y))
        {
            if (y > 0)
                sb.Append($"\\posy{y}");
            else
                sb.Append($"\\posnegy{y}");
        }
        if (fp.AnchorLock != null && ((!fp.AnchorLock.HasValue) || fp.AnchorLock.Value))
        {
            sb.Append(@"\abslock1");
        }
        else
        {
            sb.Append(@"\abslock0");
        }
        if (fp.VerticalSpace?.Value != null && int.TryParse(fp.HorizontalSpace?.Value, out int v))
        {
            sb.Append($"\\dfrmtxty{v}");
        }
        if (fp.Wrap != null && fp.Wrap.HasValue)
        {
            if (fp.Wrap.Value == TextWrappingValues.Around)
            {
                sb.Append(@"\wraparound");
            }
            else if (fp.Wrap.Value == TextWrappingValues.Through)
            {
                sb.Append(@"\wrapthrough");
            }
            else if (fp.Wrap.Value == TextWrappingValues.Tight)
            {
                sb.Append(@"\wraptight");
            }
            else if (fp.Wrap.Value == TextWrappingValues.Auto)
            {
                sb.Append(@"\wrapdefault");
            }
            else if (fp.Wrap.Value == TextWrappingValues.None)
            {
                sb.Append(@"\nowrap");
            }
            //else if (fp.Wrap.Value == TextWrappingValues.NotBeside)
            //{
            //}
        }
        if (fp.DropCap != null && fp.DropCap == DropCapLocationValues.Drop)
        {
            sb.Append(@"\dropcapt1");
        }
        else if (fp.DropCap != null && fp.DropCap == DropCapLocationValues.Margin)
        {
            sb.Append(@"\dropcapt2");
        }
        if (fp.Lines != null && fp.Lines.HasValue)
        {
            sb.Append($"\\dropcapli{fp.Lines.Value}");
        }
    }
}
