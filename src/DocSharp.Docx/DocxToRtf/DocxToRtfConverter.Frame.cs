using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal void ProcessFrameProperties(FrameProperties fp, RtfStringWriter sb)
    {
        if (fp.Width?.Value != null && int.TryParse(fp.Width.Value, out int w))
        {
            sb.Write($"\\absw{w}");
        }
        if (fp.HeightType != null)
        {
            if (fp.HeightType.Value == HeightRuleValues.Auto)
            {
                sb.Write("\\absh0");
            }
            else if (fp.Height != null && fp.Height.HasValue)
            {
                if (fp.HeightType.Value == HeightRuleValues.AtLeast)
                {
                    sb.Write($"\\absh{fp.Height.Value}");
                }
                else
                {
                    sb.Write($"\\absh-{fp.Height.Value}");
                }
            }
        }
        if (fp.HorizontalPosition?.Value != null)
        {
            if (fp.HorizontalPosition.Value == HorizontalAnchorValues.Margin)
            {
                sb.Write(@"\phmrg");
            }
            else if (fp.HorizontalPosition.Value == HorizontalAnchorValues.Page)
            {
                sb.Write(@"\phpg");
            }
            else if (fp.HorizontalPosition.Value == HorizontalAnchorValues.Text)
            {
                sb.Write(@"\phcol");
            }
        }
        if (fp.XAlign?.Value != null)
        {
            if (fp.XAlign.Value == HorizontalAlignmentValues.Center)
            {
                sb.Write("\\posxc");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Inside)
            {
                sb.Write("\\posxi");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Outside)
            {
                sb.Write("\\posxo");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Left)
            {
                sb.Write("\\posxl");
            }
            else if (fp.XAlign.Value == HorizontalAlignmentValues.Right)
            {
                sb.Write("\\posxr");
            }           
        }
        if (fp.X?.Value != null && int.TryParse(fp.X.Value, out int x))
        {
            if (x > 0)
                sb.Write($"\\posx{x}");
            else
                sb.Write($"\\posnegx{x}");
        }
        if (fp.HorizontalSpace?.Value != null && int.TryParse(fp.HorizontalSpace?.Value, out int h))
        {
            sb.Write($"\\dfrmtxtx{h}");
        }
        if (fp.VerticalPosition?.Value != null)
        {
            if (fp.VerticalPosition.Value == VerticalAnchorValues.Margin)
            {
                sb.Write(@"\pvmrg");
            }
            else if (fp.VerticalPosition.Value == VerticalAnchorValues.Page)
            {
                sb.Write(@"\pvpg");
            }
            else if (fp.VerticalPosition.Value == VerticalAnchorValues.Text)
            {
                sb.Write(@"\pvpara");
            }
        }
        if (fp.YAlign?.Value != null)
        {
            if (fp.YAlign.Value == VerticalAlignmentValues.Bottom)
            {
                sb.Write("\\posyb");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Center)
            {
                sb.Write("\\posyc");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Inline)
            {
                sb.Write("\\posyil");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Inside)
            {
                sb.Write("\\posyin");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Outside)
            {
                sb.Write("\\posyout");
            }
            else if (fp.YAlign.Value == VerticalAlignmentValues.Top)
            {
                sb.Write("\\posyt");
            }
        }
        if (fp.Y?.Value != null && int.TryParse(fp.Y.Value, out int y))
        {
            if (y > 0)
                sb.Write($"\\posy{y}");
            else
                sb.Write($"\\posnegy{y}");
        }
        if (fp.AnchorLock != null && ((!fp.AnchorLock.HasValue) || fp.AnchorLock.Value))
        {
            sb.Write(@"\abslock1");
        }
        else
        {
            sb.Write(@"\abslock0");
        }
        if (fp.VerticalSpace?.Value != null && int.TryParse(fp.HorizontalSpace?.Value, out int v))
        {
            sb.Write($"\\dfrmtxty{v}");
        }
        if (fp.Wrap != null && fp.Wrap.HasValue)
        {
            if (fp.Wrap.Value == TextWrappingValues.Around)
            {
                sb.Write(@"\wraparound");
            }
            else if (fp.Wrap.Value == TextWrappingValues.Through)
            {
                sb.Write(@"\wrapthrough");
            }
            else if (fp.Wrap.Value == TextWrappingValues.Tight)
            {
                sb.Write(@"\wraptight");
            }
            else if (fp.Wrap.Value == TextWrappingValues.Auto)
            {
                sb.Write(@"\wrapdefault");
            }
            else if (fp.Wrap.Value == TextWrappingValues.None)
            {
                sb.Write(@"\nowrap");
            }
            //else if (fp.Wrap.Value == TextWrappingValues.NotBeside)
            //{
            //}
        }
        if (fp.DropCap != null && fp.DropCap == DropCapLocationValues.Drop)
        {
            sb.Write(@"\dropcapt1");
        }
        else if (fp.DropCap != null && fp.DropCap == DropCapLocationValues.Margin)
        {
            sb.Write(@"\dropcapt2");
        }
        if (fp.Lines != null && fp.Lines.HasValue)
        {
            sb.Write($"\\dropcapli{fp.Lines.Value}");
        }
    }
}
