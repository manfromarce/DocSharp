using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;
using M = DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
	private void ProcessMathDestination(RtfDestination destination)
	{
		// Map math RTF destinations to OMML structures.
		// moMathPara -> M.Paragraph (oMathPara) containing one or more moMath
		// moMath -> M.OfficeMath
		// Process inner moMath content with a single recursive dispatcher.
		var moMathPara = destination.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "moMathPara", StringComparison.OrdinalIgnoreCase));
		if (moMathPara != null)
		{
			EnsureParagraph();
			var mathParagraph = new M.Paragraph();
			foreach (var child in moMathPara.Tokens.OfType<RtfDestination>())
			{
				if (string.Equals(child.Name, "moMath", StringComparison.OrdinalIgnoreCase))
				{
					var om = ConvertMoMath(child);
					if (om != null) mathParagraph.Append(om);
				}
			}
			if (mathParagraph.HasChildren)
				pendingParagraph!.Append(mathParagraph);
			return;
		}

		var moMath = destination.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "moMath", StringComparison.OrdinalIgnoreCase));
		if (moMath != null)
		{
			var om = ConvertMoMath(moMath);
			if (om != null)
			{
				EnsureParagraph();
				pendingParagraph!.Append(om);
			}
			return;
		}

		// Intentionally do not execute fallback image extraction here — leave empty when unsupported.
	}

	private M.OfficeMath? ConvertMoMath(RtfDestination moMathDest)
	{
		var om = new M.OfficeMath();
		ProcessMoMathChildren(moMathDest, om);
		return om;
	}

	private void ProcessMoMathChildren(RtfGroup group, OpenXmlElement parent)
	{
		if (group == null) return;

		foreach (var token in group.Tokens)
		{
			if (token is RtfDestination d)
			{				
				var name = (d.Name ?? string.Empty).ToLowerInvariant();
				OpenXmlElement? el = null;
				switch (name)
				{
					case "macc":
						el = ConvertAccent(d);
						break;
					case "mbar":
						el = ConvertBar(d);
						break;
					case "mborderbox":
						el = ConvertBorderBox(d);
						break;
					case "mbox":
						el = ConvertBox(d);
						break;
					case "md":
						el = ConvertDelimiter(d);
						break;
					case "meqarr":
						el = ConvertEquationArray(d);
						break;
					case "mf":
						el = ConvertFraction(d);
						break;
					case "mfunc":
						el = ConvertFunction(d);
						break;
					case "mgroupchr":
						el = ConvertGroupChar(d);
						break;
					case "mlimlow":
						el = ConvertLimitLower(d);
						break;
					case "mlimupp":
						el = ConvertLimitUpper(d);
						break;
					case "mm":
						el = ConvertMatrix(d);
						break;
					case "mnary":
						el = ConvertNary(d);
						break;
					case "mphant":
						el = ConvertPhantom(d);
						break;
					case "mrad":
						el = ConvertRadical(d);
						break;
					case "mspre":
						el = ConvertPreSubSuper(d);
						break;
					case "mssub":
						el = ConvertSubscript(d);
						break;
					case "mssup":
						el = ConvertSuperscript(d);
						break;
					case "mssubsup":
						el = ConvertSubSuperscript(d);
						break;
					case "mr":
						el = ConvertMathRun(d);
						break;
				}
				
				if (el != null)
					parent.Append(el);
			}
			else if (token is RtfGroup g)
			{
				// Process regular runs/chars inside math context, and recurse to find other math elements (if any)
				ConvertGroupAsMath(g, parent);
			}
			else
			{
				M.Run? run = null;

				// Use StringBuilder to handle unicode skips and surrogate pairs in the math context too.
				var sb = new StringBuilder();
				if (token is RtfControlWord cw && cw.Name.Equals("u", StringComparison.OrdinalIgnoreCase) && cw.HasValue)
				{
					ProcessUnicode(cw.Value!.Value, TryPeek(fmtStack), sb);
					run = ConvertMathRun(sb.ToString(), null, null, null, null);
				}
				else if (token is RtfText text)
				{
					HandleText(text.Text, sb);
					run = ConvertMathRun(sb.ToString(), null, null, null, null);
				}				
				else if (token is RtfChar ch)
				{
                    var enc = ResolveEncodingForRun(TryPeek(fmtStack));
					string s = enc.GetString(new byte[] { ch.CharCode });
					HandleText(s, sb);
					run = ConvertMathRun(sb.ToString(), null, null, null, null);
				}

				if (run != null)
				{
					parent.Append(run);
				}
			}		
		}
	}

	private void ConvertGroupAsMath(RtfGroup g, OpenXmlElement parent)
	{
		// Push a new state to the stack (similar to regular groups logic)
        fmtStack.Push(TryPeek(fmtStack).Clone());        
		var runState = TryPeek(fmtStack);

		// Handle relevant formatting control words
		// (does not handle all formatting of regular runs on purpose, as not everything works in math runs)
		runState.FontIndex = ReadIntegerMathProperty(g, "f");
		runState.FontSize = ReadIntegerMathProperty(g, "fs");
		runState.Bold = ReadBooleanMathProperty(g, "b") ?? false;
		runState.Italic = ReadBooleanMathProperty(g, "i") ?? false;
		runState.Underline = ReadBooleanMathProperty(g, "ul") == true ? UnderlineValues.Single : UnderlineValues.None;
		runState.Strike = ReadBooleanMathProperty(g, "strike") ?? false;
		runState.FontColorIndex = ReadIntegerMathProperty(g, "cf");
		runState.HighlightColorIndex = ReadIntegerMathProperty(g, "highlight");

		// Recurse
		ProcessMoMathChildren(g, parent);

		// Remove state from the stack
        TryPop(fmtStack);
	}

    private void ProcessArgumentProperties<T>(RtfDestination dest, T argument) where T : OfficeMathArgumentType
    {
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "margPr", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
        	argument.ArgumentProperties = new ArgumentProperties();
			var argSz = ReadIntegerMathProperty(prDest, "margSz");
			if (argSz != null)
			{
				argument.ArgumentProperties.ArgumentSize = new M.ArgumentSize() { Val = argSz.Value };
			}
		}
    }

	private M.Accent? ConvertAccent(RtfDestination dest)
	{
		var acc = new M.Accent();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "maccPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			acc.AccentProperties = new M.AccentProperties();
			var accentChar = ReadStringMathProperty(prDest, "mchr");
			if (accentChar != null && accentChar.Length == 1)
			{
				acc.AccentProperties.AccentChar = new M.AccentChar() { Val = accentChar };
			}
		}
		if (baseDest != null)
		{
			acc.Base = new M.Base();
			ProcessArgumentProperties(baseDest, acc.Base);
			ProcessMoMathChildren(baseDest, acc.Base);
		}
		return acc;
	}

    private M.Bar? ConvertBar(RtfDestination dest)
	{
		var bar = new M.Bar();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mbarPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			bar.BarProperties = new M.BarProperties();
			var pos = ReadStringMathProperty(prDest, "mpos");
			if (pos != null)
			{
				if (pos.Equals("top", StringComparison.OrdinalIgnoreCase))
					bar.BarProperties.Position = new M.Position() { Val = M.VerticalJustificationValues.Top };
				else if (pos.Equals("bot", StringComparison.OrdinalIgnoreCase))
					bar.BarProperties.Position = new M.Position() { Val = M.VerticalJustificationValues.Bottom };
			}
		}
		if (baseDest != null)
		{
			bar.Base = new M.Base();
			ProcessArgumentProperties(baseDest, bar.Base);
			ProcessMoMathChildren(baseDest, bar.Base);
		}
		return bar;
	}

	private M.BorderBox? ConvertBorderBox(RtfDestination dest)
	{
		var borderBox = new M.BorderBox();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mborderBoxPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			borderBox.BorderBoxProperties = new M.BorderBoxProperties();
			var hideBot = ReadBooleanMathProperty(prDest, "mhideBot");
			if (hideBot != null)
			{
				borderBox.BorderBoxProperties.HideBottom = new M.HideBottom() { Val = hideBot.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var hideLeft = ReadBooleanMathProperty(prDest, "mhideLeft");
			if (hideLeft != null)
			{
				borderBox.BorderBoxProperties.HideLeft = new M.HideLeft() { Val = hideLeft.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var hideRight = ReadBooleanMathProperty(prDest, "mhideRight");
			if (hideRight != null)
			{
				borderBox.BorderBoxProperties.HideRight = new M.HideRight() { Val = hideRight.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var hideTop = ReadBooleanMathProperty(prDest, "mhideTop");
			if (hideTop != null)
			{
				borderBox.BorderBoxProperties.HideTop = new M.HideTop() { Val = hideTop.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var strikeH = ReadBooleanMathProperty(prDest, "mstrikeH");
			if (strikeH != null)
			{
				borderBox.BorderBoxProperties.StrikeHorizontal = new M.StrikeHorizontal() { Val = strikeH.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var strikeV = ReadBooleanMathProperty(prDest, "mstrikeV");
			if (strikeV != null)
			{
				borderBox.BorderBoxProperties.StrikeVertical = new M.StrikeVertical() { Val = strikeV.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var strikeBLTR = ReadBooleanMathProperty(prDest, "mstrikeBLTR");
			if (strikeBLTR != null)
			{
				borderBox.BorderBoxProperties.StrikeBottomLeftToTopRight = new M.StrikeBottomLeftToTopRight() { Val = strikeBLTR.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var strikeTLBR = ReadBooleanMathProperty(prDest, "mstrikeTLBR");
			if (strikeTLBR != null)
			{
				borderBox.BorderBoxProperties.StrikeTopLeftToBottomRight = new M.StrikeTopLeftToBottomRight() { Val = strikeTLBR.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
		}
		if (baseDest != null)
		{
			borderBox.Base = new M.Base();
			ProcessArgumentProperties(baseDest, borderBox.Base);
			ProcessMoMathChildren(baseDest, borderBox.Base);
		}
		return borderBox;
	}

	private M.Box? ConvertBox(RtfDestination dest)
	{
		var box = new M.Box();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mboxPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			box.BoxProperties = new M.BoxProperties();
			var align = ReadBooleanMathProperty(prDest, "maln");
			if (align != null)
			{
				box.BoxProperties.Alignment = new M.Alignment() { Val = align.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var diff = ReadBooleanMathProperty(prDest, "mdiff");
			if (diff != null)
			{
				box.BoxProperties.Differential = new M.Differential() { Val = diff.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var noBreak = ReadBooleanMathProperty(prDest, "mnoBreak");
			if (noBreak != null)
			{
				box.BoxProperties.NoBreak = new M.NoBreak() { Val = noBreak.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var opEmu = ReadBooleanMathProperty(prDest, "mopEmu");
			if (opEmu != null)
			{
				box.BoxProperties.OperatorEmulator = new M.OperatorEmulator() { Val = opEmu.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var brk = ReadIntegerMathProperty(prDest, "mbrk");
			if (brk != null)
			{
				box.BoxProperties.Break = new M.Break() { Val = brk.Value };
			}
		}
		if (baseDest != null)
		{
			box.Base = new M.Base();
			ProcessArgumentProperties(baseDest, box.Base);
			ProcessMoMathChildren(baseDest, box.Base);
		}
		return box;
	}

	private M.Delimiter? ConvertDelimiter(RtfDestination dest)
	{
		var del = new M.Delimiter();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mdPr", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			del.DelimiterProperties = new M.DelimiterProperties();
			var beginChar = ReadStringMathProperty(prDest, "mbegChr");
			if (beginChar != null && beginChar.Length == 1)
			{
				del.DelimiterProperties.BeginChar = new M.BeginChar() { Val = beginChar };
			}
			var endChar = ReadStringMathProperty(prDest, "mendChr");
			if (endChar != null && endChar.Length == 1)
			{
				del.DelimiterProperties.EndChar = new M.EndChar() { Val = endChar };
			}
			var sepChar = ReadStringMathProperty(prDest, "msepChr");
			if (sepChar != null && sepChar.Length == 1)
			{
				del.DelimiterProperties.SeparatorChar = new M.SeparatorChar() { Val = sepChar };
			}
			var grow = ReadBooleanMathProperty(prDest, "mgrow");
			if (grow != null)
			{
				del.DelimiterProperties.GrowOperators = new M.GrowOperators() { Val = grow.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var shp = ReadStringMathProperty(prDest, "mshp");
			if (shp != null)
			{
				if (shp.Equals("match", StringComparison.OrdinalIgnoreCase))
					del.DelimiterProperties.Shape = new M.Shape() { Val = M.ShapeDelimiterValues.Match };
				else if (shp.Equals("centered", StringComparison.OrdinalIgnoreCase))
					del.DelimiterProperties.Shape = new M.Shape() { Val = M.ShapeDelimiterValues.Centered };
			}
		}
		foreach (var baseDest in dest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase)))
		{
			var e = del.AppendChild(new M.Base());
			ProcessArgumentProperties(baseDest, e);
			ProcessMoMathChildren(baseDest, e);
		}
		return del;
	}

	private M.EquationArray? ConvertEquationArray(RtfDestination dest)
	{
		var eqArr = new M.EquationArray();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "meqArrPr", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			eqArr.EquationArrayProperties = new M.EquationArrayProperties();
			var baseJc = ReadStringMathProperty(prDest, "mbaseJc");
			if (baseJc != null)
			{
				if (baseJc.Equals("top", StringComparison.OrdinalIgnoreCase))
					eqArr.EquationArrayProperties.BaseJustification = new M.BaseJustification() { Val = M.VerticalAlignmentValues.Top };
				else if (baseJc.Equals("bot", StringComparison.OrdinalIgnoreCase))
					eqArr.EquationArrayProperties.BaseJustification = new M.BaseJustification() { Val = M.VerticalAlignmentValues.Bot };
					// eqArr.EquationArrayProperties.BaseJustification = new M.BaseJustification() { Val = M.VerticalAlignmentValues.Bottom };
			}
			var maxDist = ReadBooleanMathProperty(prDest, "mmaxDist");
			if (maxDist != null)
			{
				eqArr.EquationArrayProperties.MaxDistribution = new M.MaxDistribution() { Val = maxDist.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var objDist = ReadBooleanMathProperty(prDest, "mobjDist");
			if (maxDist != null)
			{
				eqArr.EquationArrayProperties.ObjectDistribution = new M.ObjectDistribution() { Val = maxDist.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var rSp = ReadIntegerMathProperty(prDest, "mrSp");
			if (rSp != null)
			{
				eqArr.EquationArrayProperties.RowSpacing = new M.RowSpacing() { Val = rSp.Value.ToUshort() };
			}
			var rSpRule = ReadIntegerMathProperty(prDest, "mrSpRule");
			if (rSpRule != null)
			{
				eqArr.EquationArrayProperties.RowSpacingRule = new M.RowSpacingRule() { Val = rSpRule.Value };
			}
		}
		foreach (var baseDest in dest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase)))
		{
			var e = eqArr.AppendChild(new M.Base());
			ProcessArgumentProperties(baseDest, e);
			ProcessMoMathChildren(baseDest, e);
		}
		return eqArr;
	}

	private M.Fraction? ConvertFraction(RtfDestination dest)
	{
		var fraction = new M.Fraction();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mfPr", StringComparison.OrdinalIgnoreCase));
		var numDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mnum", StringComparison.OrdinalIgnoreCase));
		var denDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mden", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			fraction.FractionProperties = new M.FractionProperties();
			var type = ReadStringMathProperty(prDest, "mtype");
			if (type != null)
			{
				if (type.Equals("bar", StringComparison.OrdinalIgnoreCase))
					fraction.FractionProperties.FractionType = new M.FractionType() { Val = M.FractionTypeValues.Bar };
				else if (type.Equals("lin", StringComparison.OrdinalIgnoreCase))
					fraction.FractionProperties.FractionType = new M.FractionType() { Val = M.FractionTypeValues.Linear };
				else if (type.Equals("nobar", StringComparison.OrdinalIgnoreCase))
					fraction.FractionProperties.FractionType = new M.FractionType() { Val = M.FractionTypeValues.NoBar };
				else if (type.Equals("skw", StringComparison.OrdinalIgnoreCase))
					fraction.FractionProperties.FractionType = new M.FractionType() { Val = M.FractionTypeValues.Skewed };
			}
		}
		if (numDest != null)
		{
			fraction.Numerator = new M.Numerator();
			ProcessArgumentProperties(numDest, fraction.Numerator);
			ProcessMoMathChildren(numDest, fraction.Numerator);
		}
		if (denDest != null)
		{
			fraction.Denominator = new M.Denominator();
			ProcessArgumentProperties(denDest, fraction.Denominator);
			ProcessMoMathChildren(denDest, fraction.Denominator);
		}
		return fraction;
	}

	private M.MathFunction? ConvertFunction(RtfDestination dest)
	{
		var mathFunction = new M.MathFunction();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mfuncPr", StringComparison.OrdinalIgnoreCase));
		var fnameDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mfName", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			// mathFunction.FunctionProperties = new M.FunctionProperties();
			// Can only contain character format
		}
		if (fnameDest != null)
		{
			mathFunction.FunctionName = new M.FunctionName();
			ProcessArgumentProperties(fnameDest, mathFunction.FunctionName);
			ProcessMoMathChildren(fnameDest, mathFunction.FunctionName);
		}
		if (baseDest != null)
		{
			mathFunction.Base = new M.Base();
			ProcessArgumentProperties(baseDest, mathFunction.Base);
			ProcessMoMathChildren(baseDest, mathFunction.Base);
		}
		return mathFunction;
	}

	private M.GroupChar? ConvertGroupChar(RtfDestination dest)
	{
		var groupChar = new M.GroupChar();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mgroupChrPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			groupChar.GroupCharProperties = new M.GroupCharProperties();
			var groupCharChr = ReadStringMathProperty(prDest, "mchr");
			if (groupCharChr != null && groupCharChr.Length == 1)
			{
				groupChar.GroupCharProperties.AccentChar = new M.AccentChar() { Val = groupCharChr };
			}
			var pos = ReadStringMathProperty(prDest, "mpos");
			if (pos != null)
			{
				if (pos.Equals("top", StringComparison.OrdinalIgnoreCase))
					groupChar.GroupCharProperties.Position = new M.Position() { Val = M.VerticalJustificationValues.Top };
				else if (pos.Equals("bot", StringComparison.OrdinalIgnoreCase))
					groupChar.GroupCharProperties.Position = new M.Position() { Val = M.VerticalJustificationValues.Bottom };
			}
			var vertJc = ReadStringMathProperty(prDest, "mvertJc");
			if (vertJc != null)
			{
				if (vertJc.Equals("top", StringComparison.OrdinalIgnoreCase))
					groupChar.GroupCharProperties.VerticalJustification = new M.VerticalJustification() { Val = M.VerticalJustificationValues.Top };
				else if (vertJc.Equals("bot", StringComparison.OrdinalIgnoreCase))
					groupChar.GroupCharProperties.VerticalJustification = new M.VerticalJustification() { Val = M.VerticalJustificationValues.Bottom };
			}
		}		
		if (baseDest != null)
		{
			groupChar.Base = new M.Base();
			ProcessArgumentProperties(baseDest, groupChar.Base);
			ProcessMoMathChildren(baseDest, groupChar.Base);
		}
		return groupChar;
	}

	private M.LimitLower? ConvertLimitLower(RtfDestination dest)
	{
		var limitLow = new M.LimitLower();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mlimLowPr", StringComparison.OrdinalIgnoreCase));
		var limDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mlim", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			// limitLow.LimitLowerProperties = new M.LimitLowerProperties();
			// Can only contain character format
		}
		if (limDest != null)
		{
			limitLow.Limit = new M.Limit();
			ProcessArgumentProperties(limDest, limitLow.Limit);
			ProcessMoMathChildren(limDest, limitLow.Limit);
		}
		if (baseDest != null)
		{
			limitLow.Base = new M.Base();
			ProcessArgumentProperties(baseDest, limitLow.Base);
			ProcessMoMathChildren(baseDest, limitLow.Base);
		}
		return limitLow;
	}

	private M.LimitUpper? ConvertLimitUpper(RtfDestination dest)
	{
		var limitUpp = new M.LimitUpper();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mlimUppPr", StringComparison.OrdinalIgnoreCase));
		var limDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mlim", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			// limitUpp.LimitUpperProperties = new M.LimitUpperProperties();
			// Can only contain character format
		}
		if (limDest != null)
		{
			limitUpp.Limit = new M.Limit();
			ProcessArgumentProperties(limDest, limitUpp.Limit);
			ProcessMoMathChildren(limDest, limitUpp.Limit);
		}
		if (baseDest != null)
		{
			limitUpp.Base = new M.Base();
			ProcessArgumentProperties(baseDest, limitUpp.Base);
			ProcessMoMathChildren(baseDest, limitUpp.Base);
		}
		return limitUpp;
	}

	private M.Matrix? ConvertMatrix(RtfDestination dest)
	{
		var matrix = new M.Matrix();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mmPr", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			matrix.MatrixProperties = new M.MatrixProperties();
			var plcHide = ReadBooleanMathProperty(prDest, "mplcHide");
			if (plcHide != null)
			{
				matrix.MatrixProperties.HidePlaceholder = new M.HidePlaceholder() { Val = plcHide.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var baseJc = ReadStringMathProperty(prDest, "mbaseJc");
			if (baseJc != null)
			{
				if (baseJc.Equals("top", StringComparison.OrdinalIgnoreCase))
					matrix.MatrixProperties.BaseJustification = new M.BaseJustification() { Val = M.VerticalAlignmentValues.Top };
				else if (baseJc.Equals("bot", StringComparison.OrdinalIgnoreCase))
					matrix.MatrixProperties.BaseJustification = new M.BaseJustification() { Val = M.VerticalAlignmentValues.Bot };
					// matrix.MatrixProperties.BaseJustification = new M.BaseJustification() { Val = M.VerticalAlignmentValues.Bottom };
			}
			var cGp = ReadIntegerMathProperty(prDest, "mcGp");
			if (cGp != null)
			{
				matrix.MatrixProperties.ColumnGap = new M.ColumnGap() { Val = cGp.Value.ToUshort() };
			}
			var cGpRule = ReadIntegerMathProperty(prDest, "mcGpRule");
			if (cGpRule != null)
			{
				matrix.MatrixProperties.ColumnGapRule = new M.ColumnGapRule() { Val = cGpRule.Value };
			}
			var cSp = ReadIntegerMathProperty(prDest, "mcSp");
			if (cSp != null)
			{
				matrix.MatrixProperties.ColumnSpacing = new M.ColumnSpacing() { Val = cSp.Value.ToUshort() };
			}
			var rSp = ReadIntegerMathProperty(prDest, "mrSp");
			if (rSp != null)
			{
				matrix.MatrixProperties.RowSpacing = new M.RowSpacing() { Val = rSp.Value.ToUshort() };
			}
			var rSpRule = ReadIntegerMathProperty(prDest, "mrSpRule");
			if (rSpRule != null)
			{
				matrix.MatrixProperties.RowSpacingRule = new M.RowSpacingRule() { Val = rSpRule.Value };
			}
			var matrixColumnsDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name.Equals("mmcs", StringComparison.OrdinalIgnoreCase));
			if (matrixColumnsDest != null)
			{
				matrix.MatrixProperties.MatrixColumns = new M.MatrixColumns();
				foreach (var columnDest in matrixColumnsDest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "mmc", StringComparison.OrdinalIgnoreCase)))
				{
					var column = matrix.MatrixProperties.MatrixColumns.AppendChild(new M.MatrixColumn());			
					var columnPrDest = columnDest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => d.Name.Equals("mmcPr", StringComparison.OrdinalIgnoreCase));
					if (columnPrDest != null)
					{
						var columnPr = column.AppendChild(new M.MatrixColumnProperties());
						var mjc = ReadStringMathProperty(prDest, "mmjc");
						if (mjc != null)
						{
							if (mjc.Equals("left", StringComparison.OrdinalIgnoreCase))
								columnPr.MatrixColumnJustification = new M.MatrixColumnJustification() { Val = M.HorizontalAlignmentValues.Left };
							else if (mjc.Equals("center", StringComparison.OrdinalIgnoreCase))
								columnPr.MatrixColumnJustification = new M.MatrixColumnJustification() { Val = M.HorizontalAlignmentValues.Center };
							else if (mjc.Equals("right", StringComparison.OrdinalIgnoreCase))
								columnPr.MatrixColumnJustification = new M.MatrixColumnJustification() { Val = M.HorizontalAlignmentValues.Right };
						}
						var count = ReadIntegerMathProperty(prDest, "count");
						if (count != null)
						{
							columnPr.MatrixColumnCount = new M.MatrixColumnCount() { Val = count.Value };
						}
					}
				}
			}
		}
		foreach (var matrixRowDest in dest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "mmr", StringComparison.OrdinalIgnoreCase)))
		{
			var mr = matrix.AppendChild(new M.MatrixRow());			
			foreach (var matrixRowBaseDest in matrixRowDest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase)))
			{
				var me = mr.AppendChild(new M.Base());
				ProcessArgumentProperties(matrixRowBaseDest, me);
				ProcessMoMathChildren(matrixRowBaseDest, me);
			}			
		}
		return matrix;
	}

	private M.Nary? ConvertNary(RtfDestination dest)
	{
		var nary = new M.Nary();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mnaryPr", StringComparison.OrdinalIgnoreCase));
		var subDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msub", StringComparison.OrdinalIgnoreCase));
		var supDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msup", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			nary.NaryProperties = new M.NaryProperties();
			var naryChar = ReadStringMathProperty(prDest, "mchr");
			if (naryChar != null && naryChar.Length == 1)
			{
				nary.NaryProperties.AccentChar = new M.AccentChar() { Val = naryChar };
			}
			var grow = ReadBooleanMathProperty(prDest, "mgrow");
			if (grow != null)
			{
				nary.NaryProperties.GrowOperators = new M.GrowOperators() { Val = grow.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var limloc = ReadStringMathProperty(prDest, "mlimLoc");
			if (limloc != null)
			{
				if (limloc.Equals("undovr", StringComparison.OrdinalIgnoreCase))
					nary.NaryProperties.LimitLocation = new M.LimitLocation() { Val = M.LimitLocationValues.UnderOver };
				else if (limloc.Equals("subsup", StringComparison.OrdinalIgnoreCase))
					nary.NaryProperties.LimitLocation = new M.LimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
			}
			var subHide = ReadBooleanMathProperty(prDest, "msubHide");
			if (subHide != null)
			{
				nary.NaryProperties.HideSubArgument = new M.HideSubArgument() { Val = subHide.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var supHide = ReadBooleanMathProperty(prDest, "msupHide");
			if (supHide != null)
			{
				nary.NaryProperties.HideSuperArgument = new M.HideSuperArgument() { Val = supHide.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
		}
		if (subDest != null)
		{
			nary.SubArgument = new M.SubArgument();
			ProcessArgumentProperties(subDest, nary.SubArgument);
			ProcessMoMathChildren(subDest, nary.SubArgument);
		}
		if (supDest != null)
		{
			nary.SuperArgument = new M.SuperArgument();
			ProcessArgumentProperties(supDest, nary.SuperArgument);
			ProcessMoMathChildren(supDest, nary.SuperArgument);
		}
		if (baseDest != null)
		{
			nary.Base = new M.Base();
			ProcessArgumentProperties(baseDest, nary.Base);
			ProcessMoMathChildren(baseDest, nary.Base);
		}
		return nary;
	}

	private M.Phantom? ConvertPhantom(RtfDestination dest)
	{
		var phant = new M.Phantom();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mphantPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			phant.PhantomProperties = new M.PhantomProperties();
			var show = ReadBooleanMathProperty(prDest, "mshow");
			if (show != null)
			{
				phant.PhantomProperties.ShowPhantom = new M.ShowPhantom() { Val = show.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var transp = ReadBooleanMathProperty(prDest, "mtransp");
			if (transp != null)
			{
				phant.PhantomProperties.Transparent = new M.Transparent() { Val = transp.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var zeroAsc = ReadBooleanMathProperty(prDest, "mzeroAsc");
			if (zeroAsc != null)
			{
				phant.PhantomProperties.ZeroAscent = new M.ZeroAscent() { Val = zeroAsc.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var zeroDesc = ReadBooleanMathProperty(prDest, "mzeroDesc");
			if (zeroDesc != null)
			{
				phant.PhantomProperties.ZeroDescent = new M.ZeroDescent() { Val = zeroDesc.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
			var zeroWid = ReadBooleanMathProperty(prDest, "mzeroWid");
			if (zeroWid != null)
			{
				phant.PhantomProperties.ZeroWidth = new M.ZeroWidth() { Val = zeroWid.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
		}
		if (baseDest != null)
		{
			phant.Base = new M.Base();
			ProcessArgumentProperties(baseDest, phant.Base);
			ProcessMoMathChildren(baseDest, phant.Base);
		}
		return phant;
	}

	private M.Radical? ConvertRadical(RtfDestination dest)
	{
		var rad = new M.Radical();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mradPr", StringComparison.OrdinalIgnoreCase));
		var degreeDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mdeg", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			rad.RadicalProperties = new M.RadicalProperties();
			var degHide = ReadBooleanMathProperty(prDest, "mdegHide");
			if (degHide != null)
			{
				rad.RadicalProperties.HideDegree = new M.HideDegree() { Val = degHide.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}			
		}
		if (baseDest != null)
		{
			rad.Base = new M.Base();
			ProcessArgumentProperties(baseDest, rad.Base);
			ProcessMoMathChildren(baseDest, rad.Base);
		}
		if (degreeDest != null)
		{
			rad.Degree = new M.Degree();
			ProcessArgumentProperties(degreeDest, rad.Degree);
			ProcessMoMathChildren(degreeDest, rad.Degree);
		}
		return rad;
	}

	private M.PreSubSuper? ConvertPreSubSuper(RtfDestination dest)
	{
		var preSubSup = new M.PreSubSuper();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msPrePr", StringComparison.OrdinalIgnoreCase));
		var subDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msub", StringComparison.OrdinalIgnoreCase));
		var supDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msup", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			// preSubSup.PreSubSuperProperties = new M.PreSubSuperProperties();
			// Can only contain character format
		}
		if (subDest != null)
		{
			preSubSup.SubArgument = new M.SubArgument();
			ProcessArgumentProperties(subDest, preSubSup.SubArgument);
			ProcessMoMathChildren(subDest, preSubSup.SubArgument);
		}
		if (supDest != null)
		{
			preSubSup.SuperArgument = new M.SuperArgument();
			ProcessArgumentProperties(supDest, preSubSup.SuperArgument);
			ProcessMoMathChildren(supDest, preSubSup.SuperArgument);
		}
		if (baseDest != null)
		{
			preSubSup.Base = new M.Base();
			ProcessArgumentProperties(baseDest, preSubSup.Base);
			ProcessMoMathChildren(baseDest, preSubSup.Base);
		}
		return preSubSup;
	}

	private M.SubSuperscript? ConvertSubSuperscript(RtfDestination dest)
	{
		var subSup = new M.SubSuperscript();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msSubSupPr", StringComparison.OrdinalIgnoreCase));
		var subDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msub", StringComparison.OrdinalIgnoreCase));
		var supDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msup", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			subSup.SubSuperscriptProperties = new M.SubSuperscriptProperties();
			var alnScr = ReadBooleanMathProperty(prDest, "malnScr");
			if (alnScr != null)
			{
				subSup.SubSuperscriptProperties.AlignScripts = new M.AlignScripts() { Val = alnScr.Value ? M.BooleanValues.On : M.BooleanValues.Off };
			}
		}
		if (subDest != null)
		{
			subSup.SubArgument = new M.SubArgument();
			ProcessArgumentProperties(subDest, subSup.SubArgument);
			ProcessMoMathChildren(subDest, subSup.SubArgument);
		}
		if (supDest != null)
		{
			subSup.SuperArgument = new M.SuperArgument();
			ProcessArgumentProperties(supDest, subSup.SuperArgument);
			ProcessMoMathChildren(supDest, subSup.SuperArgument);
		}
		if (baseDest != null)
		{
			subSup.Base = new M.Base();
			ProcessArgumentProperties(baseDest, subSup.Base);
			ProcessMoMathChildren(baseDest, subSup.Base);
		}
		return subSup;
	}

	private M.Subscript? ConvertSubscript(RtfDestination dest)
	{
		var sub = new M.Subscript();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msSubPr", StringComparison.OrdinalIgnoreCase));
		var subDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msub", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			// sub.SubscriptProperties = new M.SubscriptProperties();
			// Can only contain character format
		}
		if (subDest != null)
		{
			sub.SubArgument = new M.SubArgument();
			ProcessArgumentProperties(subDest, sub.SubArgument);
			ProcessMoMathChildren(subDest, sub.SubArgument);
		}
		if (baseDest != null)
		{
			sub.Base = new M.Base();
			ProcessArgumentProperties(baseDest, sub.Base);
			ProcessMoMathChildren(baseDest, sub.Base);
		}
		return sub;
	}

	private M.Superscript? ConvertSuperscript(RtfDestination dest)
	{
		var sup = new M.Superscript();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msSupPr", StringComparison.OrdinalIgnoreCase));
		var supDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msup", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			// sup.SuperscriptProperties = new M.SuperscriptProperties();
			// Can only contain character format
		}
		if (supDest != null)
		{
			sup.SuperArgument = new M.SuperArgument();
			ProcessArgumentProperties(supDest, sup.SuperArgument);
			ProcessMoMathChildren(supDest, sup.SuperArgument);
		}
		if (baseDest != null)
		{
			sup.Base = new M.Base();
			ProcessArgumentProperties(baseDest, sup.Base);
			ProcessMoMathChildren(baseDest, sup.Base);
		}
		return sup;
	}

	// RTF uses two ways to assign property values for math, depending on the property: 
	// - the standard RTF way with a parameter N, as in \msty2
	// - a mini group like {\mtype skw}
	private bool? ReadBooleanMathProperty(RtfGroup parentGroupOrDest, string propertyName)
	{
		// This function detects the first
		// - \propertyNameN or {\propertyNameN} where N is 0 (false) or 1 (true). If N is not present, assume true.
		// - or {\propertyName X}, where X is 0, 1, on or off.

		if (parentGroupOrDest == null) return null;

		// 1) Top-level control word: \property or \propertyN
		var cw = parentGroupOrDest.Tokens.OfType<RtfControlWord>().FirstOrDefault(c => c.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
		if (cw != null)
		{
			if (cw.HasValue)
			{
				return cw.Value == 1;
			}
			// control word present without explicit value -> assume true
			return true;
		}

		// 2) Subgroup or destination: {\propertyN} or {\property x}
		var group = parentGroupOrDest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, propertyName, StringComparison.OrdinalIgnoreCase)) ?? 
		            parentGroupOrDest.Tokens.OfType<RtfGroup>().FirstOrDefault(g => g.Tokens.FirstOrDefault() is RtfControlWord cw && cw.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
		if (group != null)
		{
			var innerCw = group.Tokens.FirstOrDefault() as RtfControlWord;
			if (innerCw != null)
			{
				if (innerCw.HasValue) // control word nested in brackets: {\propertyN}
				{
					// 0 = false, 1 = true
					return innerCw.Value == 1;
				}
				else // possible group of type {\property X}		
				{
					// The value should be 0, 1, on or off without escaped chars or similar, unless the RTF is not conformant.
					var text = parentGroupOrDest.Tokens.FirstOrDefault() as RtfText;
					if (text != null && !string.IsNullOrWhiteSpace(text.Text))
					{
						if (text.Text.Equals("0", StringComparison.OrdinalIgnoreCase) || text.Text.Equals("off", StringComparison.OrdinalIgnoreCase))
							return false;
						else if (text.Text.Equals("1", StringComparison.OrdinalIgnoreCase) || text.Text.Equals("on", StringComparison.OrdinalIgnoreCase))
							return true;
					}
					else 
					{
						// property is present without a value -> assume true
						return true;
					}
				}
			}
			else // possible destination of type {\property X}, in this case the control word name is "eaten" by the destination.
			// (we don't consider {\destN}, because the only one that can be produced by RtfReader is \pnseclvlN, but it is not relevant in this context)
			{				
				var text = parentGroupOrDest.Tokens.FirstOrDefault() as RtfText;
				if (text != null && !string.IsNullOrWhiteSpace(text.Text))
				{
					if (text.Text.Equals("0", StringComparison.OrdinalIgnoreCase) || text.Text.Equals("off", StringComparison.OrdinalIgnoreCase))
						return false;
					else if (text.Text.Equals("1", StringComparison.OrdinalIgnoreCase) || text.Text.Equals("on", StringComparison.OrdinalIgnoreCase))
						return true;
				}
				else 
				{
					// property is present without a value -> assume true
					return true;
				}
			}
		}
		return null;
	}

	private int? ReadIntegerMathProperty(RtfGroup parentGroupOrDest, string propertyName)
	{
		// This function detects the first \propertyNameN, {\propertyNameN} or {\propertyName N}, where N is a number.

		if (parentGroupOrDest == null) return null;

		// 1) Top-level controlword: \propertyN
		var cw = parentGroupOrDest.Tokens.OfType<RtfControlWord>().FirstOrDefault(c => c.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
		if (cw != null)
		{
			if (cw.HasValue)
			{
				return cw.Value;
			}
			// control word present without numeric value -> no integer
			return null;
		}

		// 2) Subgroup or destination: {\propertyN} or {\property x}
		var group = parentGroupOrDest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, propertyName, StringComparison.OrdinalIgnoreCase)) ?? 
				    parentGroupOrDest.Tokens.OfType<RtfGroup>().FirstOrDefault(g => g.Tokens.FirstOrDefault() is RtfControlWord cw && cw.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
		if (group != null)
		{
			var innerCw = group.Tokens.FirstOrDefault() as RtfControlWord;
			if (innerCw != null)
			{
				if (innerCw.HasValue) // control word nested in brackets: {\propertyN}
				{
					return innerCw.Value;
				}
				else // possible group of type {\property N}		
				{
					

					var text = parentGroupOrDest.Tokens.FirstOrDefault() as RtfText; // The number cannot be an escaped char or similar, unless the RTF is not conformant
					if (text != null && !string.IsNullOrEmpty(text.Text) && int.TryParse(text.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
						return n;
				}
			}
			else // possible destination of type {\property N}, in this case the control word name is "eaten" by the destination.
			// (we don't consider {\destN}, because the only one that can be produced by RtfReader is \pnseclvlN, but it is not relevant in this context)
			{				
				var text = parentGroupOrDest.Tokens.FirstOrDefault() as RtfText; // The number cannot be an escaped char or similar, unless the RTF is not conformant
				if (text != null && !string.IsNullOrEmpty(text.Text) && int.TryParse(text.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
					return n;
			}
		}

		return null;
	}

	private string? ReadStringMathProperty(RtfGroup parentGroupOrDest, string propertyName)
	{
		// This function detects the first {\propertyName x} where x is a string, escaped char, Unicode char or combination of these.

		if (parentGroupOrDest == null) return null;

		// 1) Subgroup or destination: {\property x}
		var group = parentGroupOrDest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, propertyName, StringComparison.OrdinalIgnoreCase)) ?? 
				    parentGroupOrDest.Tokens.OfType<RtfGroup>().FirstOrDefault(g => g.Tokens.OfType<RtfControlWord>().Any(c => c.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase)));
		if (group != null)
		{
			var sb = new StringBuilder();
			ConvertGroupAsText(group, sb);
			// If subgroup exists but is empty, returns an empty string (it may be acceptable depending on the property).
			return sb.ToString();
		}

		// 2) Top-level controlword with value: \propertyN -> return N as string
		// var cw = parentGroupOrDest.Tokens.OfType<RtfControlWord>().FirstOrDefault(c => c.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
		// if (cw != null)
		// {
		// 	if (cw.HasValue) return cw.Value.ToString();
		// 	// Present without a value -> treat as empty string
		// 	return string.Empty;
		// }

		return null;
	}

	private M.Run? ConvertMathRun(RtfDestination dest)
	{
		var mnor = ReadBooleanMathProperty(dest, "mnor");
		var mlit = ReadBooleanMathProperty(dest, "mlit");
		var msty = ReadIntegerMathProperty(dest, "msty");
		var mscr = ReadIntegerMathProperty(dest, "mscr");

		var builder = new StringBuilder();
		ConvertGroupAsText(dest, builder);
		return ConvertMathRun(builder.ToString(), msty, mnor, mlit, mscr);
	}

	private M.Run? ConvertMathRun(string text, int? style = null, bool? normal = null, bool? literal = null, int? scr = null)
	{
		var run = new M.Run
        {
			RunProperties = CreateRunProperties(TryPeek(fmtStack)),
            MathRunProperties = new M.RunProperties()
        };
		// Font family is ignored on purpose for now, because fonts such as "Cambria Math Greek" are frequently specified in RTF 
		// but they cause rendering issues in DOCX; other fonts do not work at all in math runs.
		// However, we use them to retrieve the encoding from the font table.
		run.RunProperties.RunFonts = new RunFonts() { Ascii = "Cambria Math", HighAnsi = "Cambria Math", EastAsia = "Cambria Math", ComplexScript = "Cambria Math" };

		if (normal != null)
			run.MathRunProperties.AppendChild(new M.NormalText() { Val = normal.Value ? M.BooleanValues.On : M.BooleanValues.Off });

		if (literal != null)
			run.MathRunProperties.AppendChild(new M.Literal() { Val = literal.Value ? M.BooleanValues.On : M.BooleanValues.Off });

		if (style != null)
		{
			switch (style.Value)
			{
				case 0:
					run.MathRunProperties.AppendChild(new M.Style() { Val = M.StyleValues.Plain });
					break;
				case 1:
					run.MathRunProperties.AppendChild(new M.Style() { Val = M.StyleValues.Bold });
					break;
				case 2:
					run.MathRunProperties.AppendChild(new M.Style() { Val = M.StyleValues.Italic });
					break;
				case 3:
					run.MathRunProperties.AppendChild(new M.Style() { Val = M.StyleValues.BoldItalic });
					break;
			}
		}

		if (scr != null)
		{
			switch (scr.Value)
			{
				case 0:
					run.MathRunProperties.AppendChild(new M.Script() { Val = M.ScriptValues.Roman });
					break;
				case 1:
					run.MathRunProperties.AppendChild(new M.Script() { Val = M.ScriptValues.Script });
					break;
				case 2:
					run.MathRunProperties.AppendChild(new M.Script() { Val = M.ScriptValues.Fraktur });
					break;
				case 3:
					run.MathRunProperties.AppendChild(new M.Script() { Val = M.ScriptValues.DoubleStruck });
					break;
				case 4:
					run.MathRunProperties.AppendChild(new M.Script() { Val = M.ScriptValues.SansSerif });
					break;
				case 5:
					run.MathRunProperties.AppendChild(new M.Script() { Val = M.ScriptValues.Script });
					break;
			}
		}
		run.AppendChild(new M.Text() { Text = text, Space = SpaceProcessingModeValues.Preserve });
		return run;
	}
}