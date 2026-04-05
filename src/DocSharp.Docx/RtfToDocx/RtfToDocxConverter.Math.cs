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

	private void ProcessMoMathChildren(RtfGroup group, OpenXmlElement parent, string fontFamily = "Cambria Math")
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
						el = ConvertMathRun(d, fontFamily);
						break;
					default:
						// TODO: better handle regular runs/chars inside math context
						// For now just recurse to find \mr
						ProcessMoMathChildren(d, parent);
						break;
				}
				if (el != null)
				{
					parent.Append(el);
				}
			}
			else if (token is RtfGroup g)
			{
				// TODO: better handle regular runs/chars inside math context.
				// For now just recurse to find \mr and keep track of custom font (if specified)
				var fontCw = g.Tokens.OfType<RtfControlWord>().FirstOrDefault(x => x.Name.Equals("f", StringComparison.OrdinalIgnoreCase));
				if (fontCw != null && fontCw.HasValue && fontTable.TryGetValue(fontCw.Value!.Value, out var finfo) && !string.IsNullOrEmpty(finfo?.Name))
				{
					ProcessMoMathChildren(g, parent, finfo.Name);
				}
				else
				{
					ProcessMoMathChildren(g, parent);
				}
			}
		}
	}

	private void ProcessMathProperties(RtfGroup group, OpenXmlElement parent)
	{
		if (group == null) return;
		foreach (var d in group.Tokens.OfType<RtfDestination>())
		{
			// TODO
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
			ProcessMathProperties(prDest, acc.AccentProperties);
		}
		if (baseDest != null)
		{
			acc.Base = new M.Base();
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
			ProcessMathProperties(prDest, bar.BarProperties);
		}
		if (baseDest != null)
		{
			bar.Base = new M.Base();
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
			ProcessMathProperties(prDest, borderBox.BorderBoxProperties);
		}
		if (baseDest != null)
		{
			borderBox.Base = new M.Base();
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
			ProcessMathProperties(prDest, box.BoxProperties);
		}
		if (baseDest != null)
		{
			box.Base = new M.Base();
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
			ProcessMathProperties(prDest, del.DelimiterProperties);
		}
		foreach (var runDest in dest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase)))
		{
			var e = del.AppendChild(new M.Base());
			ProcessMoMathChildren(runDest, e);
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
			ProcessMathProperties(prDest, eqArr.EquationArrayProperties);
		}
		foreach (var runDest in dest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase)))
		{
			var e = eqArr.AppendChild(new M.Base());
			ProcessMoMathChildren(runDest, e);
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
			ProcessMathProperties(prDest, fraction.FractionProperties);
		}
		if (numDest != null)
		{
			fraction.Numerator = new M.Numerator();
			ProcessMoMathChildren(numDest, fraction.Numerator);
		}
		if (denDest != null)
		{
			fraction.Denominator = new M.Denominator();
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
			mathFunction.FunctionProperties = new M.FunctionProperties();
			ProcessMathProperties(prDest, mathFunction.FunctionProperties);
		}
		if (fnameDest != null)
		{
			mathFunction.FunctionName = new M.FunctionName();
			ProcessMoMathChildren(fnameDest, mathFunction.FunctionName);
		}
		if (baseDest != null)
		{
			mathFunction.Base = new M.Base();
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
			ProcessMathProperties(prDest, groupChar.GroupCharProperties);
		}		
		if (baseDest != null)
		{
			groupChar.Base = new M.Base();
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
			limitLow.LimitLowerProperties = new M.LimitLowerProperties();
			ProcessMathProperties(prDest, limitLow.LimitLowerProperties);
		}
		if (limDest != null)
		{
			limitLow.Limit = new M.Limit();
			ProcessMoMathChildren(limDest, limitLow.Limit);
		}
		if (baseDest != null)
		{
			limitLow.Base = new M.Base();
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
			limitUpp.LimitUpperProperties = new M.LimitUpperProperties();
			ProcessMathProperties(prDest, limitUpp.LimitUpperProperties);
		}
		if (limDest != null)
		{
			limitUpp.Limit = new M.Limit();
			ProcessMoMathChildren(limDest, limitUpp.Limit);
		}
		if (baseDest != null)
		{
			limitUpp.Base = new M.Base();
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
			ProcessMathProperties(prDest, matrix.MatrixProperties);
		}
		foreach (var matrixRowDest in dest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "mmr", StringComparison.OrdinalIgnoreCase)))
		{
			var mr = matrix.AppendChild(new M.MatrixRow());			
			foreach (var matrixRowBaseDest in matrixRowDest.Tokens.OfType<RtfDestination>().Where(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase)))
			{
				var me = mr.AppendChild(new M.Base());
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
			ProcessMathProperties(prDest, nary.NaryProperties);
		}
		if (subDest != null)
		{
			nary.SubArgument = new M.SubArgument();
			ProcessMoMathChildren(subDest, nary.SubArgument);
		}
		if (supDest != null)
		{
			nary.SuperArgument = new M.SuperArgument();
			ProcessMoMathChildren(supDest, nary.SuperArgument);
		}
		if (baseDest != null)
		{
			nary.Base = new M.Base();
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
			ProcessMathProperties(prDest, phant.PhantomProperties);
		}
		if (baseDest != null)
		{
			phant.Base = new M.Base();
			ProcessMoMathChildren(baseDest, phant.Base);
		}
		return phant;
	}

	private M.Radical? ConvertRadical(RtfDestination dest)
	{
		var rad = new M.Radical();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "mradPr", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			rad.RadicalProperties = new M.RadicalProperties();
			ProcessMathProperties(prDest, rad.RadicalProperties);
		}
		if (baseDest != null)
		{
			rad.Base = new M.Base();
			ProcessMoMathChildren(baseDest, rad.Base);
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
			preSubSup.PreSubSuperProperties = new M.PreSubSuperProperties();
			ProcessMathProperties(prDest, preSubSup.PreSubSuperProperties);
		}
		if (subDest != null)
		{
			preSubSup.SubArgument = new M.SubArgument();
			ProcessMoMathChildren(subDest, preSubSup.SubArgument);
		}
		if (supDest != null)
		{
			preSubSup.SuperArgument = new M.SuperArgument();
			ProcessMoMathChildren(supDest, preSubSup.SuperArgument);
		}
		if (baseDest != null)
		{
			preSubSup.Base = new M.Base();
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
			ProcessMathProperties(prDest, subSup.SubSuperscriptProperties);
		}
		if (subDest != null)
		{
			subSup.SubArgument = new M.SubArgument();
			ProcessMoMathChildren(subDest, subSup.SubArgument);
		}
		if (supDest != null)
		{
			subSup.SuperArgument = new M.SuperArgument();
			ProcessMoMathChildren(supDest, subSup.SuperArgument);
		}
		if (baseDest != null)
		{
			subSup.Base = new M.Base();
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
			sub.SubscriptProperties = new M.SubscriptProperties();
			ProcessMathProperties(prDest, sub.SubscriptProperties);
		}
		if (subDest != null)
		{
			sub.SubArgument = new M.SubArgument();
			ProcessMoMathChildren(subDest, sub.SubArgument);
		}
		if (baseDest != null)
		{
			sub.Base = new M.Base();
			ProcessMoMathChildren(baseDest, sub.Base);
		}
		return sub;
	}

	private M.Superscript? ConvertSuperscript(RtfDestination dest)
	{
		var sup = new M.Superscript();
		var prDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msSubPr", StringComparison.OrdinalIgnoreCase));
		var supDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "msup", StringComparison.OrdinalIgnoreCase));
		var baseDest = dest.Tokens.OfType<RtfDestination>().FirstOrDefault(d => string.Equals(d.Name, "me", StringComparison.OrdinalIgnoreCase));
		if (prDest != null)
		{
			sup.SuperscriptProperties = new M.SuperscriptProperties();
			ProcessMathProperties(prDest, sup.SuperscriptProperties);
		}
		if (supDest != null)
		{
			sup.SuperArgument = new M.SuperArgument();
			ProcessMoMathChildren(supDest, sup.SuperArgument);
		}
		if (baseDest != null)
		{
			sup.Base = new M.Base();
			ProcessMoMathChildren(baseDest, sup.Base);
		}
		return sup;
	}

	private M.Run? ConvertMathRun(RtfDestination dest, string fontFamily = "Cambria Math")
	{
		var run = new M.Run();
        run.RunProperties = new RunProperties
        {
			// RunFonts = new RunFonts() { Ascii = fontFamily, HighAnsi = fontFamily, EastAsia = fontFamily, ComplexScript = fontFamily }
            RunFonts = new RunFonts() { Ascii = "Cambria Math", HighAnsi = "Cambria Math", EastAsia = "Cambria Math", ComplexScript = "Cambria Math" }
            // For now font family is ignored, because fonts such as "Cambria Math Greek" are frequently specified in RTF 
			// but they cause rendering issues in DOCX. 
			// However, we use them to detect the encoding.
        };
		run.MathRunProperties = new M.RunProperties();
		var msty = dest.Tokens.OfType<RtfControlWord>().FirstOrDefault(p => p.Name.Equals("msty", StringComparison.OrdinalIgnoreCase)) ?? 
				   dest.Tokens.OfType<RtfGroup>().FirstOrDefault(g => g.Tokens.FirstOrDefault() is RtfControlWord cw && cw.Name.Equals("msty", StringComparison.OrdinalIgnoreCase))?.Tokens.First() as RtfControlWord;
		var mnor = dest.Tokens.OfType<RtfControlWord>().FirstOrDefault(p => p.Name.Equals("mnor", StringComparison.OrdinalIgnoreCase)); 
		var mlit = dest.Tokens.OfType<RtfControlWord>().FirstOrDefault(p => p.Name.Equals("mlit", StringComparison.OrdinalIgnoreCase)); 		
		
		if (mnor != null)
		{
			if (!mnor.HasValue || mnor.Value == 1)
				run.MathRunProperties.AppendChild(new M.NormalText() { Val = M.BooleanValues.On });
		}
		if (mlit != null)
		{
			if (!mlit.HasValue || mlit.Value == 1)
				run.MathRunProperties.Literal = new M.Literal() { Val = M.BooleanValues.On };
		}
		if (msty != null)
		{
			if (msty.HasValue)
			{				
				switch(msty.Value!.Value)
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
		}
		var builder = new StringBuilder();
		// Try to determine an encoding from the math run font family (if present).
		Encoding? overrideEnc = null;
		if (!string.IsNullOrEmpty(fontFamily))
		{
			var finfo = fontTable.Values.FirstOrDefault(fi => string.Equals(fi.Name, fontFamily, StringComparison.OrdinalIgnoreCase));
			if (finfo != null)
				overrideEnc = GetEncodingFromFontInfo(finfo);
		}
		ConvertGroupAsText(dest, builder, overrideEnc);
		run.AppendChild(new M.Text() { Text = builder.ToString(), Space = SpaceProcessingModeValues.Preserve });
		return run;
	}

}