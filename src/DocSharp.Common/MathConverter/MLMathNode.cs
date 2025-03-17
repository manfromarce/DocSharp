using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using DocSharp.Xml;
#if NETFRAMEWORK
using DocSharp.Helpers;
#endif

namespace DocSharp.MathConverter;

// Convert oMath element of omml to latex. Named 'oMath2Latex' in Python version.
internal class MLMathNode : MLNodeBase
{
    private static readonly HashSet<string> direct_tags = 
        ["box", "sSub", "sSup", "sSubSup", "sPre", "num", "den", "deg", "e", "phant", "borderBox"];

    public string Text { get; }

    public MLMathNode(XmlNode oMathNode)
    {
        //mxd. Sanity check...
        if (oMathNode.Name != "m:oMath")
            throw new InvalidDataException($"Expected 'oMath' node, but got '{oMathNode.LocalName}'!");

        Text = ProcessChildren(oMathNode);
        // Remove unnecessary spaces before right curly brackets
        string pattern = @"(?<!\\)\s+\}";
        Text = Regex.Replace(Text, pattern, "}");
    }

    protected override TeXNode? ProcessTag(string tag, XmlNode elm)
    {
        switch (tag)
        {
            case "acc": return DoAcc(elm);
            case "r": return DoR(elm);
            case "bar": return DoBar(elm);
            case "sub": return DoSub(elm);
            case "sup": return DoSup(elm);
            case "f": return DoF(elm);
            case "func": return DoFunc(elm);
            case "fName": return DoFName(elm);
            case "groupChr": return DoGroupChr(elm);
            case "d": return DoD(elm);
            case "rad": return DoRad(elm);
            case "eqArr": return DoEqArr(elm);
            case "limLow": return DoLimLow(elm);
            case "limUpp": return DoLimUpp(elm);
            case "lim": return DoLim(elm);
            case "m": return DoM(elm);
            case "mr": return DoMr(elm);
            case "nary": return DoNary(elm);
            //case "borderBox": return DoBorderBox(elm); // currently only the content is preserved (without border)
            default: return null;
        }
    }

    protected override TeXNode? ProcessUnknown(XmlNode elm, string tag)
    {
        if (direct_tags.Contains(tag))
            return new TeXNode(ProcessChildren(elm));

        if (tag.EndsWith("Pr"))
            return new TeXNode(new MLPropertiesNode(elm));

        return null;
    }

    // The Accent function
    private TeXNode? DoAcc(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        var pr = children["accPr"];
        var latex_str = GetValue(pr.GetAttributeValue("chr"), CHR_DEFAULT["ACC_VAL"], CHR);

        if (string.IsNullOrEmpty(latex_str))
            return null;
        else
            return new TeXNode(string.Format(latex_str, children["e"]));
    }

    // The Bar function
    private TeXNode? DoBar(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        var pr = children["barPr"];
        var latex_str = GetValue(pr.GetAttributeValue("pos"), POS_DEFAULT["BAR_VAL"], POS);

        if (string.IsNullOrEmpty(latex_str))
            return null;
        else
            return new TeXNode(pr + string.Format(latex_str, children["e"]));
    }

    // The Delimiter object
    private TeXNode DoD(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        var pr = children["dPr"];

        var null_val = D_DEFAULT["null"];
        var beg_chr = GetValue(pr.GetAttributeValue("begChr"), D_DEFAULT["left"], T);
        var end_chr = GetValue(pr.GetAttributeValue("endChr"), D_DEFAULT["right"], T);

        var left = string.IsNullOrEmpty(beg_chr) ? null_val : EscapeLatex(beg_chr);
        var right = string.IsNullOrEmpty(end_chr) ? null_val : EscapeLatex(end_chr);
        var text = children["e"].ToString();

        return new TeXNode(pr + string.Format(D, left, text, right));
    }

    // The Substript object
    private TeXNode? DoSub(XmlNode elm)
    {
        var val = ProcessChildren(elm);
        return (val == null ? null : new TeXNode(string.Format(SUB, val)));
    }

    // The Superstript object
    private TeXNode? DoSup(XmlNode elm)
    {
        var val = ProcessChildren(elm);
        return (val == null ? null : new TeXNode(string.Format(SUP, val)));
    }

    // The Fraction object
    private TeXNode? DoF(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        var pr = children["fPr"];
        var frac = GetValue(pr.GetAttributeValue("type"), F_DEFAULT, F);

        if (string.IsNullOrEmpty(frac))
            return null;
        else
            return new TeXNode(pr + string.Format(frac, children["num"], children["den"]));
    }

    // The Function-Apply object (examples: sin cos)
    private TeXNode? DoFunc(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        if (children.ContainsKey("fName"))
        {
            var func_name = children["fName"].ToString();
            return new TeXNode(func_name.Replace(FUNC_PLACE, children["e"].ToString()));
        }
        else
        {
            // Unsupported function name
            return null;
        }
    }

    // The func name
    private TeXNode? DoFName(XmlNode elm)
    {
        var latex_func_names = new List<string>();
        foreach (var info in ProcessChildrenList(elm))
        {
            var fkey = info.Node.ToString();

            if (info.Tag == "r")
            {
                if (FUNC.TryGetValue(fkey, out string? value))
                    latex_func_names.Add(value);
                else
                    // Function names may be localized in Office math,
                    // use the function name directly as fallback.
                    latex_func_names.Add($" {fkey}({{fe}})"); 
            }
            else
            {
                latex_func_names.Add(fkey);
            }
        }

        var func_names = string.Join("", latex_func_names);
        return new TeXNode(func_names.Contains(FUNC_PLACE) ? func_names : func_names + FUNC_PLACE);
    }

    // The Group-Character object
    private TeXNode? DoGroupChr(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        var pr = children["groupChrPr"];
        var chr = GetValue(pr.GetAttributeValue("chr"));

        if (string.IsNullOrEmpty(chr))
            return null;
        else
            return new TeXNode(pr + string.Format(chr, children["e"]));
    }

    // The radical object
    private TeXNode DoRad(XmlNode elm)
    {
        var children = ProcessChildrenDict(elm);
        if (children.TryGetValue("deg", out TeXNode? value) && !string.IsNullOrEmpty(value.ToString()))
            return new TeXNode(string.Format(RAD, value, children["e"]));

        return new TeXNode(string.Format(RAD_DEFAULT, children["e"]));
    }

    // The Array object
    private TeXNode DoEqArr(XmlNode elm)
    {
        var include = new HashSet<string> { "e" };
        var text = string.Join(BRK, ProcessChildrenList(elm, include).Select(t => t.Node));        
        return new TeXNode(string.Format(ARR, text));
    }

    // The Lower-Limit object
    private TeXNode DoLimLow(XmlNode elm)
    {
        var include = new HashSet<string> { "e", "lim" };
        var children = ProcessChildrenDict(elm, include);
        var func_name = children["e"].ToString();

        if (!LIM_FUNC.ContainsKey(func_name))
            throw new NotSupportedException($"Not supported lim function: '{func_name}'!");

        return new TeXNode(string.Format(LIM_FUNC[func_name], children["lim"]));
    }

    // The Upper-Limit object
    private TeXNode DoLimUpp(XmlNode elm)
    {
        var include = new HashSet<string> { "e", "lim" };
        var children = ProcessChildrenDict(elm, include);

        return new TeXNode(string.Format(LIM_UPP, children["lim"], children["e"]));
    }

    // The lower limit of the limLow object and the upper limit of the limUpp function
    private TeXNode DoLim(XmlNode elm)
    {
        var result = ProcessChildren(elm).Replace(LIM_TO[0], LIM_TO[1]);
        return new TeXNode(result);
    }

    // The Matrix object
    private TeXNode DoM(XmlNode elm)
    {
        var rows = new List<string>();
        foreach (var info in ProcessChildrenList(elm))
            if (info.Tag == "mr")
                rows.Add(info.Node.ToString());

        return new TeXNode(string.Format(M, string.Join(BRK, rows)));
    }

    // A single row of the matrix m
    private TeXNode DoMr(XmlNode elm)
    {
        var include = new HashSet<string> { "e" };
        var result = string.Join("&", ProcessChildrenList(elm, include).Select(t => t.Node));

        return new TeXNode(result);
    }

    // The n-ary object
    private TeXNode? DoNary(XmlNode elm)
    {
        var res = new List<string>();
        string? bo = "";

        foreach (var info in ProcessChildrenList(elm))
        {
            if (info.Tag == "naryPr")
            {
                bo = GetValue(info.Node.GetAttributeValue("chr"), store: CHR_BO);

                //TODO: mxd. Very unsure, but MS Word seems to assume this as default
                if (bo == null)
                    bo = "\\int";
            }
            else if (info.Tag == "e" && IsComplexEquation(info.Node.ToString())) //mxd. Wrap equation in {} to better match MSWord LaTeX output. 
            {
                res.Add($"{{{info.Node}}}");
            }
            else
            {
                res.Add($"{info.Node}");
            }
        }

        //mxd. Handle operators with empty <m:sub> / <m:sup> nodes (like '\sum x')
        var val = string.Join("", res);
        if (!val.StartsWith('_') && !val.StartsWith('^'))
            bo += " ";

        return new TeXNode(bo + val);
    }

    //private TeXNode? DoBorderBox(XmlNode elm)
    //{
    //    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math.borderbox?view=openxml-3.0.1
    //    return null;
    //}

    // Get text from 'r' element and try convert them to latex symbols
    private TeXNode? DoR(XmlNode elm)
    {
        var t = elm.GetChildByName("m:t");
        if (t == null)
            return null;

        bool keepSpaces = (t.GetAttributeValue("xml:space") == "preserve");
        return new TeXNode(EscapeLatex(t.InnerText, keepSpaces));
    }

    private static string EscapeLatex(string str, bool keepSpaces = true)
    {
        // This function must only be called for leaf values 
        // coming from Office math direcly, so that backslash and other special chars
        // are meant to be part of the math expression itself rather than LaTex syntax.
        var sb = new StringBuilder();

        foreach (char c in str)
        {
            if (CharsToEscape.Contains(c))
                sb.Append(@"\" + c);
            else if (c == '~')
                sb.Append(@"\sim ");
                //sb.Append(@"\text{\textasciitilde}");
            else if (c == '^')
                sb.Append(@"\text{\textasciicircum}");
            else if (c == '\\')
                sb.Append(@"\backslash ");
            else if (keepSpaces && c == ' ')
                sb.Append(@"\ ");
            else
                sb.Append(T.TryGetValue(c.ToString(), out string? value) ? value : c);
        }

        return sb.ToString();
    }

    private static string? GetValue(string? key, string? default_value = null, Dictionary<string, string>? store = null)
    {
        if (key == null) 
            return default_value;
            // Don't treat empty string as null in this case,
            // as having attributes explicitly set to "" may have a different meaning.

        store ??= CHR;

        return store.TryGetValue(key, out string? value) ? value : key;
    }

    // Return true when equation needs to be wrapped in {}.
    private static bool IsComplexEquation(string e)
    {
        return e.Length > 2 && (e.Contains('+') || e.Contains('-') || e.Contains('\\'));
    }

}
