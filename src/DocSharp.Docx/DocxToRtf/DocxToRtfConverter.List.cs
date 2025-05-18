using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessNumberingPart(Numbering numbering, StringBuilder sb)
    {
        sb.Append(@"{\*\listtable");
        ProcessListTable(numbering, sb);
        sb.Append('}');
        sb.Append(@"{\*\listoverridetable");
        ProcessListOverrideTable(numbering, sb);
        sb.Append('}');
    }

    private void ProcessListTable(Numbering numbering, StringBuilder sb)
    {
        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            sb.Append(@"{\list");
            if (abstractNum.Nsid?.Val != null)
            {
                sb.Append(@$"\listid{abstractNum.Nsid.Val.ToLong()}");
            }
            if (abstractNum.TemplateCode?.Val != null)
            {
                sb.Append(@$"\listemplateid{abstractNum.TemplateCode.Val.ToLong()}");
            }
            if (abstractNum.MultiLevelType?.Val != null)
            {
                if (abstractNum.MultiLevelType.Val == MultiLevelValues.SingleLevel)
                {
                    sb.Append(@"\listsimple1");
                }
                else if (abstractNum.MultiLevelType.Val == MultiLevelValues.Multilevel)
                {
                    sb.Append(@"\listsimple0"); // default value if MultilevelType is not specified
                }
                else if (abstractNum.MultiLevelType.Val == MultiLevelValues.HybridMultilevel)
                {
                    sb.Append(@"\listhybrid");
                }
            }
            foreach (var level in abstractNum.Elements<Level>())
            {
                ProcessLevel(level, sb);
            }
            if (abstractNum.AbstractNumDefinitionName?.Val != null)
            {
                sb.Append(@$"{{\listname {abstractNum.AbstractNumDefinitionName.Val};}}");
            }
            if (abstractNum.StyleLink != null)
            {
            }
            if (abstractNum.NumberingStyleLink != null)
            {
            }
            sb.AppendLineCrLf('}');
        }

        if (numbering.Elements<NumberingPictureBullet>().Any()) 
        { 
            sb.Append(@"{\*\listpicture");
            foreach (var pictureBullet in numbering.Elements<NumberingPictureBullet>())
            {
                //if (pictureBullet.NumberingPictureBulletId != null)
                //{
                // Not available in RTF, add picture bullets in order instead
                //}
                if (pictureBullet.PictureBulletBase != null)
                {
                    ProcessPictureBulletBase(pictureBullet.PictureBulletBase, sb);
                }
                else if (pictureBullet.Drawing != null)
                {
                    ProcessDrawing(pictureBullet.Drawing, sb);
                }
                else if (pictureBullet.Elements<Picture>().FirstOrDefault() is Picture pict)
                {
                    ProcessPicture(pict, sb); // PictureBulletBase might incorrectly be interpreted as Picture
                }
                else if (pictureBullet.GetFirstChild<AlternateContent>() is AlternateContent alternateContent)
                {
                    if (alternateContent.Descendants<PictureBulletBase>().FirstOrDefault() is PictureBulletBase pbb)
                    {
                        ProcessPictureBulletBase(pbb, sb);
                    }
                    else if (alternateContent.Descendants<Drawing>().FirstOrDefault() is Drawing drawing1)
                    {
                        ProcessDrawing(drawing1, sb);
                    }
                    else if (alternateContent.Descendants<Picture>().FirstOrDefault() is Picture pict1)
                    {
                        ProcessPicture(pict1, sb);
                    }
                }
            }
            sb.AppendLineCrLf('}');
        }
    }

    private void ProcessLevel(Level level, StringBuilder sb)
    {
        sb.Append(@"{\listlevel");
        
        if (level.Tentative != null && level.Tentative.HasValue && level.Tentative.Value)
        {
            sb.Append(@"\lvltentative");
        }
        //if (level.LevelIndex != null)
        //{
        // Not supported in RTF, add levels in order instead
        //}
        if (level.LevelRestart is LevelRestart restart)
        {

        }
        if (level.LevelPictureBulletId is LevelPictureBulletId pictureBulletId)
        {
            sb.Append($@"\levelpicture{pictureBulletId.Val}");
        }
        if (level.LevelJustification?.Val != null)
        {
            if (level.LevelJustification.Val == LevelJustificationValues.Left)
            {
                sb.Append(@"\leveljcn0");
            }
            else if (level.LevelJustification.Val == LevelJustificationValues.Center)
            {
                sb.Append(@"\leveljcn1");
            }
            else if (level.LevelJustification.Val == LevelJustificationValues.Right)
            {
                sb.Append(@"\leveljcn2");
            }
        }

        if (level.LevelSuffix?.Val != null)
        {
            if (level.LevelSuffix.Val == LevelSuffixValues.Tab)
            {
                sb.Append("\\levelfollow0"); // Default
            }
            else if (level.LevelSuffix.Val == LevelSuffixValues.Space)
            {
                sb.Append("\\levelfollow1"); 
            }
            else if (level.LevelSuffix.Val == LevelSuffixValues.Nothing)
            {
                sb.Append("\\levelfollow2"); 
            }
        }

        if (level.LevelText is LevelText levelText)
        {
            string levelTemplateId = "";
            if (level.TemplateCode?.Value != null)
            {
                levelTemplateId = @$"\listemplateid{level.TemplateCode.ToLong()}";
            }
            ProcessLevelText(levelText, levelTemplateId, sb);
        }

        if (level.LegacyNumbering is LegacyNumbering legacy &&
            (legacy.Legacy == null || legacy.Legacy.Value))
        {
            sb.Append(@"\levelold1");
            if (legacy.LegacyIndent != null && int.TryParse(legacy.LegacyIndent, out int legacyIndent))
            {
                sb.Append($@"\levelindent{legacyIndent}");
            }
            if (legacy.LegacySpace != null && int.TryParse(legacy.LegacySpace, out int legacySpace))
            {
                sb.Append($@"\levelspace{legacySpace}");
            }
        }
        else
        {
            sb.Append(@"\levelindent0\levelspace0");
        }
        if (level.StartNumberingValue?.Val != null)
        {
            sb.Append($@"\levelstartat{level.StartNumberingValue.Val}");
        }

        if (level.NumberingFormat is NumberingFormat numberingFormat)
        {
            ProcessNumberingFormat(numberingFormat, sb);
        }
        // The numbering format might be specified in an <mc:Choice> or <mc:Fallback> element
        else if (level.Descendants<NumberingFormat>().FirstOrDefault() is NumberingFormat nf)
        {
            ProcessNumberingFormat(nf, sb);
        }

        if (level.IsLegalNumberingStyle != null && (level.IsLegalNumberingStyle.Val == null || level.IsLegalNumberingStyle.Val))
        {
            sb.Append($@"\levellegal1");
        }
        if (level.PreviousParagraphProperties is PreviousParagraphProperties prevParagraphProperties)
        {
            ProcessPreviousParagraphProperties(prevParagraphProperties, sb);
        }
        if (level.NumberingSymbolRunProperties is NumberingSymbolRunProperties symbolRunProperties)
        {
            ProcessNumberingSymbolRunProperties(symbolRunProperties, sb);
        }
        if (level.ParagraphStyleIdInLevel is ParagraphStyleIdInLevel paragraphStyleIdInLevel)
        {

        }

        sb.AppendLineCrLf('}');
    }

    private void ProcessNumberingSymbolRunProperties(NumberingSymbolRunProperties symbolRunProperties, StringBuilder sb)
    {
        ProcessRunFormatting(symbolRunProperties, sb);
    }

    private void ProcessPreviousParagraphProperties(PreviousParagraphProperties prevParagraphProperties, StringBuilder sb)
    {
        if (prevParagraphProperties.Indentation is Indentation ind)
        {
            if (ind?.Left != null)
            {
                sb.Append($"\\li{ind.Left}");
            }
            else if (ind?.Start != null)
            {
                sb.Append($"\\lin{ind.Start}");                 
            }

            if (ind?.FirstLine != null)
            {
                sb.Append($"\\fi{ind.FirstLine}");
            }
            else if (ind?.Hanging != null)
            {
                sb.Append($"\\fi-{ind.Hanging}");
            }
            // Others are not supported in this context in RTF.
        }
    }

    private void ProcessNumberingFormat(NumberingFormat numberingFormat, StringBuilder sb)
    {
        if (numberingFormat.Val == null)
        {
            // Default
            sb.Append(@"\levelnfc0");
            return;
        }

        if (numberingFormat.Val == NumberFormatValues.Decimal)
        {
            sb.Append(@"\levelnfc0"); // 1, 2, 3
        }
        else if (numberingFormat.Val == NumberFormatValues.UpperRoman)
        {
            sb.Append(@"\levelnfc1"); // I, II, III
        }
        else if (numberingFormat.Val == NumberFormatValues.LowerRoman)
        {
            sb.Append(@"\levelnfc2"); // i, ii, iii
        }
        else if (numberingFormat.Val == NumberFormatValues.UpperLetter)
        {
            sb.Append(@"\levelnfc3"); // A, B, C
        }
        else if (numberingFormat.Val == NumberFormatValues.LowerLetter)
        {
            sb.Append(@"\levelnfc4"); // a, b, c
        }
        else if (numberingFormat.Val == NumberFormatValues.Ordinal)
        {
            sb.Append(@"\levelnfc5"); // 1st, 2nd, 3rd
        }
        else if (numberingFormat.Val == NumberFormatValues.CardinalText)
        {
            sb.Append(@"\levelnfc6"); // One, Two Three
        }
        else if (numberingFormat.Val == NumberFormatValues.OrdinalText)
        {
            sb.Append(@"\levelnfc7"); // First, Second, Third
        }
        else if (numberingFormat.Val == NumberFormatValues.ChineseCounting ||
                 numberingFormat.Val == NumberFormatValues.IdeographDigital ||
                 numberingFormat.Val == NumberFormatValues.KoreanDigital ||
                 numberingFormat.Val == NumberFormatValues.TaiwaneseCounting)
        {
            sb.Append(@"\levelnfc10"); // Kanji numbering without the digit character (DBNUM1)
        }
        else if (numberingFormat.Val == NumberFormatValues.ChineseLegalSimplified ||
                 numberingFormat.Val == NumberFormatValues.IdeographLegalTraditional ||
                 numberingFormat.Val == NumberFormatValues.JapaneseCounting ||
                 numberingFormat.Val == NumberFormatValues.KoreanCounting)
        {
            sb.Append(@"\levelnfc11"); // Kanji numbering with the digit character (DBNUM2)
        }
        else if (numberingFormat.Val == NumberFormatValues.Aiueo)
        {
            sb.Append(@"\levelnfc12"); // 46 phonetic katakana characters in "aiueo" order (AIUEO) (newer form – “あいうえお。。。” based on phonem matrix) 
        }
        else if (numberingFormat.Val == NumberFormatValues.Iroha)
        {
            sb.Append(@"\levelnfc13"); // 46 phonetic katakana characters in "iroha" order (IROHA) (old form – “いろはにほへとちりぬるお。。。” based on haiku from long ago) 
        }
        else if (numberingFormat.Val == NumberFormatValues.ChineseCountingThousand ||
                numberingFormat.Val == NumberFormatValues.JapaneseLegal ||
                numberingFormat.Val == NumberFormatValues.KoreanLegal ||
                numberingFormat.Val == NumberFormatValues.TaiwaneseCountingThousand)
        {
            sb.Append(@"\levelnfc14"); // Kanji numbering 3 (DBNUM3)
        }
        else if (numberingFormat.Val == NumberFormatValues.KoreanDigital2 ||
                numberingFormat.Val == NumberFormatValues.TaiwaneseDigital)
        {
            sb.Append(@"\levelnfc15"); // Kanji numbering 4 (DBNUM4)
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedCircle)
        {
            sb.Append(@"\levelnfc18"); // Circle numbering (CIRCLENUM) 
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalFullWidth || 
                 numberingFormat.Val == NumberFormatValues.DecimalFullWidth2)
        {
            sb.Append(@"\levelnfc19"); // Double-byte Arabic numbering 
        }
        else if (numberingFormat.Val == NumberFormatValues.Bullet)
        {
            sb.Append(@"\levelnfc23"); // Bullet (no number)
        }
        else if (numberingFormat.Val == NumberFormatValues.Ganada)
        {
            sb.Append(@"\levelnfc24"); // Korean numbering 2 (GANADA) 
        }
        else if (numberingFormat.Val == NumberFormatValues.Chosung)
        {
            sb.Append(@"\levelnfc25"); // Korean numbering 1 (Chosung) 
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedFullstop)
        {
            sb.Append(@"\levelnfc26"); // Chinese numbering 1 (GB1)
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedParen)
        {
            sb.Append(@"\levelnfc27"); // Chinese numbering 2 (GB2)
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedCircleChinese)
        {
            sb.Append(@"\levelnfc28"); // Chinese numbering 3 (GB3)
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographEnclosedCircle)
        {
            sb.Append(@"\levelnfc29"); // Chinese numbering 4 (GB4)
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographTraditional)
        {
            sb.Append(@"\levelnfc30"); // Chinese Zodiac numbering 1 (ZODIAC1) 
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographZodiac)
        {
            sb.Append(@"\levelnfc31"); // Chinese Zodiac numbering 2 (ZODIAC2) 
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographZodiacTraditional)
        {
            sb.Append(@"\levelnfc32"); // Chinese Zodiac numbering 3 (ZODIAC3) 
        }
        else if (numberingFormat.Val == NumberFormatValues.Hebrew1)
        {
            sb.Append(@"\levelnfc45"); // Hebrew non-standard decimal 
        }
        else if (numberingFormat.Val == NumberFormatValues.ArabicAlpha)
        {
            sb.Append(@"\levelnfc46"); // Arabic Alif Ba Tah
        }
        else if (numberingFormat.Val == NumberFormatValues.Hebrew2)
        {
            sb.Append(@"\levelnfc47"); // Hebrew Biblical standard
        }
        else if (numberingFormat.Val == NumberFormatValues.ArabicAbjad)
        {
            sb.Append(@"\levelnfc48"); // Arabic Abjad style 
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiVowels)
        {
            sb.Append(@"\levelnfc49");
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiConsonants)
        {
            sb.Append(@"\levelnfc50"); 
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiNumbers)
        {
            sb.Append(@"\levelnfc51");
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiCounting)
        {
            sb.Append(@"\levelnfc52"); // Hindi descriptive (cardinals) 
        }
        else if (numberingFormat.Val == NumberFormatValues.ThaiLetters)
        {
            sb.Append(@"\levelnfc53");
        }
        else if (numberingFormat.Val == NumberFormatValues.ThaiNumbers)
        {
            sb.Append(@"\levelnfc54");
        }
        else if (numberingFormat.Val == NumberFormatValues.ThaiCounting)
        {
            sb.Append(@"\levelnfc55"); // Thai descriptive (cardinals) 
        }
        else if (numberingFormat.Val == NumberFormatValues.VietnameseCounting)
        {
            sb.Append(@"\levelnfc56"); // Vietnamese descriptive (cardinals) 
        }
        else if (numberingFormat.Val == NumberFormatValues.NumberInDash)
        {
            sb.Append(@"\levelnfc57"); // Page number format - # -
        }
        else if (numberingFormat.Val == NumberFormatValues.RussianLower)
        {
            sb.Append(@"\levelnfc58"); 
        }
        else if (numberingFormat.Val == NumberFormatValues.RussianUpper)
        {
            sb.Append(@"\levelnfc59"); 
        }
        else if (numberingFormat.Val == NumberFormatValues.Custom)
        {
            if (numberingFormat.Format != null)
            {
                switch (numberingFormat.Format.Value)
                {
                    // These are few standard formats created by Microsoft Word
                    case "01, 02, 03, ...":
                        sb.Append(@"\levelnfc22");
                        break;
                    case "001, 002, 003, ...":
                        sb.Append(@"\levelnfc62");
                        break;
                    case "0001, 0002, 0003, ...":
                        sb.Append(@"\levelnfc63");
                        break;
                    case "00001, 00002, 00003, ...":
                        sb.Append(@"\levelnfc64"); 
                        break;
                }
            }
        }
        else if (numberingFormat.Val == NumberFormatValues.None)
        {
            sb.Append(@"\levelnfc255");
        }
        else
        {
            // Default (decimal numbers)
            sb.Append(@"\levelnfc0");
        }
    }

    private void ProcessLevelText(LevelText levelText, string levelTemplateId, StringBuilder sb)
    {
        sb.Append(@"{\leveltext");
        sb.Append(levelTemplateId);
        string src = string.Empty;
        if (levelText?.Val?.Value != null)
        {
            src = levelText.Val.Value;
            var rtfParts = new List<string>();
            int i = 0;
            while (i < src.Length)
            {
                if (src[i] == '%' && i + 1 < src.Length && char.IsDigit(src[i + 1]))
                {
                    // Find number after %
                    int start = i + 1;
                    int len = 1;
                    while (start + len < src.Length && char.IsDigit(src[start + len]))
                        len++;
                    int level = int.Parse(src.Substring(start, len));
                    int rtfLevel = level - 1;
                    rtfParts.Add($@"\'{rtfLevel:X2}");
                    i = start + len;
                }
                else
                {
                    // Keep regular characters
                    rtfParts.Add(RtfHelpers.EscapeChar(src[i]));
                    i++;
                }
            }
            // Length is the number of elements in rtfParts (level numbers count as one character)
            sb.Append($@"\'{rtfParts.Count:X2}");
            foreach (var part in rtfParts)
                sb.Append(part);
        }
        sb.Append(";}");
        
        WriteLevelNumbers(src, sb);
    }

    private void WriteLevelNumbers(string src, StringBuilder sb)
    {
        // Write \levelnumbers control word followed by the level numbers offset.
        // For example:
        // 1.1.1. --> \'01\'03\'05
        // (1) --> '\02
        // bullet --> just {\levelnumbers;} (no numbers)

        var offsets = new List<int>();
        int offset = 1; // offset starts from 1 (skip length)
        int i = 0;
        while (i < src.Length)
        {
            if (src[i] == '%' && i + 1 < src.Length && char.IsDigit(src[i + 1]))
            {
                int start = i + 1;
                int len = 1;
                while (start + len < src.Length && char.IsDigit(src[start + len]))
                    len++;
                offsets.Add(offset);
                offset++; // Each \'xx counts as 1
                i = start + len;
            }
            else
            {
                // Regular chars (escaped or not) count as 1
                offset++;
                i++;
            }
        }

        sb.Append("{\\levelnumbers");
        foreach (var pos in offsets)
        {
            sb.Append($"\\'{pos:X2}");
        }
        sb.Append(";}");
    }

    private void ProcessListOverrideTable(Numbering numbering, StringBuilder sb)
    {
        foreach (var num in numbering.Elements<NumberingInstance>())
        {
            sb.Append(@"{\listoverride"); 

            // Get list id from the AbstractNum element
            if (num.AbstractNumId?.Val != null &&
                numbering.Elements<AbstractNum>().FirstOrDefault(x => x.AbstractNumberId != null && 
                                                                      x.AbstractNumberId == num.AbstractNumId.Val) 
                                                  is AbstractNum abstractNum && 
                abstractNum.Nsid?.Val != null)
            {
                sb.Append(@$"\listid{abstractNum.Nsid.Val.ToLong()}");
            }

            if (num.NumberID != null && num.NumberID.HasValue)
            {
                sb.Append(@$"\ls{num.NumberID.Value}");
            }

            var levelOverrides = num.Elements<LevelOverride>();
            if (levelOverrides == null || !levelOverrides.Any())
            {
                sb.Append(@"\listoverridecount0");
            }
            else
            {
                if (levelOverrides.Count() == 1)
                {
                    sb.Append(@"\listoverridecount1");
                }
                else
                {
                    sb.Append(@"\listoverridecount9");
                }

                foreach (var levelOverride in levelOverrides)
                {
                    sb.Append(@"{\lfolevel");
                    if (levelOverride.StartOverrideNumberingValue?.Val != null)
                    {
                        sb.Append(@$"\listoverridestartat{levelOverride.StartOverrideNumberingValue.Val} ");
                    }
                    /* TODO: if both the start-at and the format are overridden, 
                     * put the \levelstartatN inside the \listlevel contained in the \lfolevel
                     */
                    if (levelOverride.Level is Level level)
                    {
                        sb.Append(@"\listoverrideformat");
                        // TODO: 1, 9 or 0 (not added by Word ?)

                        ProcessLevel(level, sb);
                    }
                    //if (levelOverride.LevelIndex != null)
                    //{
                    // Not supported in RTF, add levels in order instead
                    //}
                    sb.Append('}');
                }
            }
            sb.AppendLineCrLf('}');
        }
    }

    internal void ProcessListItem(NumberingProperties numPr, StringBuilder sb)
    {
        if (numPr.NumberingLevelReference?.Val != null && numPr.NumberingId?.Val != null)
        {
            // TODO: Calculate list text for RTF readers that don't support automatic numbering
            // or Word 97-2007 (or newer) lists. For now, just use a generic bullet.
            fonts.TryAddAndGetIndex("Arial", out int fontIndex);
            sb.Append($@"{{\listtext \f{fontIndex}\bullet\tab}}");

            sb.Append($@"\ls{numPr.NumberingId.Val}\ilvl{numPr.NumberingLevelReference.Val}");
        }
    }
}
