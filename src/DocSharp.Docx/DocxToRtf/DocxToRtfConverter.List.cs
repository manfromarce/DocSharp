using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal void ProcessNumberingPart(Numbering numbering, RtfStringWriter sb)
    {
        sb.Write(@"{\*\listtable");
        ProcessListTable(numbering, sb);
        sb.Write('}');
        sb.Write(@"{\*\listoverridetable");
        ProcessListOverrideTable(numbering, sb);
        sb.Write('}');
    }

    private void ProcessListTable(Numbering numbering, RtfStringWriter sb)
    {
        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            sb.Write(@"{\list");
            if (abstractNum.Nsid?.Val != null)
            {
                sb.Write(@$"\listid{abstractNum.Nsid.Val.ToLong()}");
            }
            if (abstractNum.TemplateCode?.Val != null)
            {
                sb.Write(@$"\listemplateid{abstractNum.TemplateCode.Val.ToLong()}");
            }
            if (abstractNum.MultiLevelType?.Val != null)
            {
                if (abstractNum.MultiLevelType.Val == MultiLevelValues.SingleLevel)
                {
                    sb.Write(@"\listsimple1");
                }
                else if (abstractNum.MultiLevelType.Val == MultiLevelValues.Multilevel)
                {
                    sb.Write(@"\listsimple0"); // default value if MultilevelType is not specified
                }
                else if (abstractNum.MultiLevelType.Val == MultiLevelValues.HybridMultilevel)
                {
                    sb.Write(@"\listhybrid");
                }
            }
            foreach (var level in abstractNum.Elements<Level>())
            {
                ProcessLevel(level, sb);
            }
            if (abstractNum.AbstractNumDefinitionName?.Val != null)
            {
                sb.Write(@$"{{\listname {abstractNum.AbstractNumDefinitionName.Val};}}");
            }
            if (abstractNum.StyleLink != null)
            {
            }
            if (abstractNum.NumberingStyleLink != null)
            {
            }
            sb.WriteLine('}');
        }

        if (numbering.Elements<NumberingPictureBullet>().Any()) 
        { 
            sb.Write(@"{\*\listpicture");
            foreach (var pictureBullet in numbering.Elements<NumberingPictureBullet>())
            {
                //if (pictureBullet.NumberingPictureBulletId != null)
                //{
                // Not available in RTF, add picture bullets in order instead
                //}
                if (pictureBullet.PictureBulletBase != null)
                {
                    ProcessVml(pictureBullet.PictureBulletBase, sb);
                }
                else if (pictureBullet.Drawing != null)
                {
                    ProcessDrawing(pictureBullet.Drawing, sb);
                }
                else if (pictureBullet.Elements<Picture>().FirstOrDefault() is Picture pict)
                {
                    ProcessVml(pict, sb); // PictureBulletBase might incorrectly be interpreted as Picture
                }
                else if (pictureBullet.GetFirstChild<AlternateContent>() is AlternateContent alternateContent)
                {
                    if (alternateContent.Descendants<PictureBulletBase>().FirstOrDefault() is PictureBulletBase pbb)
                    {
                        ProcessVml(pbb, sb);
                    }
                    else if (alternateContent.Descendants<Drawing>().FirstOrDefault() is Drawing drawing1)
                    {
                        ProcessDrawing(drawing1, sb);
                    }
                    else if (alternateContent.Descendants<Picture>().FirstOrDefault() is Picture pict1)
                    {
                        ProcessVml(pict1, sb);
                    }
                }
            }
            sb.WriteLine('}');
        }
    }

    private void ProcessLevel(Level level, RtfStringWriter sb)
    {
        sb.Write(@"{\listlevel");
        
        if (level.Tentative != null && level.Tentative.HasValue && level.Tentative.Value)
        {
            sb.Write(@"\lvltentative");
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
            sb.Write($@"\levelpicture{pictureBulletId.Val}");
        }
        if (level.LevelJustification?.Val != null)
        {
            if (level.LevelJustification.Val == LevelJustificationValues.Left)
            {
                sb.Write(@"\leveljcn0");
            }
            else if (level.LevelJustification.Val == LevelJustificationValues.Center)
            {
                sb.Write(@"\leveljcn1");
            }
            else if (level.LevelJustification.Val == LevelJustificationValues.Right)
            {
                sb.Write(@"\leveljcn2");
            }
        }

        if (level.LevelSuffix?.Val != null)
        {
            if (level.LevelSuffix.Val == LevelSuffixValues.Tab)
            {
                sb.Write("\\levelfollow0"); // Default
            }
            else if (level.LevelSuffix.Val == LevelSuffixValues.Space)
            {
                sb.Write("\\levelfollow1"); 
            }
            else if (level.LevelSuffix.Val == LevelSuffixValues.Nothing)
            {
                sb.Write("\\levelfollow2"); 
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
            sb.Write(@"\levelold1");
            if (legacy.LegacyIndent != null && int.TryParse(legacy.LegacyIndent, out int legacyIndent))
            {
                sb.Write($@"\levelindent{legacyIndent}");
            }
            if (legacy.LegacySpace != null && int.TryParse(legacy.LegacySpace, out int legacySpace))
            {
                sb.Write($@"\levelspace{legacySpace}");
            }
        }
        else
        {
            sb.Write(@"\levelindent0\levelspace0");
        }
        if (level.StartNumberingValue?.Val != null)
        {
            sb.Write($@"\levelstartat{level.StartNumberingValue.Val}");
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
            sb.Write($@"\levellegal1");
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

        sb.WriteLine('}');
    }

    private void ProcessNumberingSymbolRunProperties(NumberingSymbolRunProperties symbolRunProperties, RtfStringWriter sb)
    {
        ProcessRunFormatting(symbolRunProperties, sb);
    }

    private void ProcessPreviousParagraphProperties(PreviousParagraphProperties prevParagraphProperties, RtfStringWriter sb)
    {
        if (prevParagraphProperties.Indentation is Indentation ind)
        {
            if (ind?.Left != null)
            {
                sb.Write($"\\li{ind.Left}");
            }
            else if (ind?.Start != null)
            {
                sb.Write($"\\lin{ind.Start}");                 
            }

            if (ind?.FirstLine != null)
            {
                sb.Write($"\\fi{ind.FirstLine}");
            }
            else if (ind?.Hanging != null)
            {
                sb.Write($"\\fi-{ind.Hanging}");
            }
            // Others are not supported in this context in RTF.
        }
    }

    private void ProcessNumberingFormat(NumberingFormat numberingFormat, RtfStringWriter sb)
    {
        if (numberingFormat.Val == null)
        {
            // Default
            sb.Write(@"\levelnfc0");
            return;
        }

        if (numberingFormat.Val == NumberFormatValues.Decimal)
        {
            sb.Write(@"\levelnfc0"); // 1, 2, 3
        }
        else if (numberingFormat.Val == NumberFormatValues.UpperRoman)
        {
            sb.Write(@"\levelnfc1"); // I, II, III
        }
        else if (numberingFormat.Val == NumberFormatValues.LowerRoman)
        {
            sb.Write(@"\levelnfc2"); // i, ii, iii
        }
        else if (numberingFormat.Val == NumberFormatValues.UpperLetter)
        {
            sb.Write(@"\levelnfc3"); // A, B, C
        }
        else if (numberingFormat.Val == NumberFormatValues.LowerLetter)
        {
            sb.Write(@"\levelnfc4"); // a, b, c
        }
        else if (numberingFormat.Val == NumberFormatValues.Ordinal)
        {
            sb.Write(@"\levelnfc5"); // 1st, 2nd, 3rd
        }
        else if (numberingFormat.Val == NumberFormatValues.CardinalText)
        {
            sb.Write(@"\levelnfc6"); // One, Two Three
        }
        else if (numberingFormat.Val == NumberFormatValues.OrdinalText)
        {
            sb.Write(@"\levelnfc7"); // First, Second, Third
        }
        else if (numberingFormat.Val == NumberFormatValues.ChineseCounting ||
                 numberingFormat.Val == NumberFormatValues.IdeographDigital ||
                 numberingFormat.Val == NumberFormatValues.KoreanDigital ||
                 numberingFormat.Val == NumberFormatValues.TaiwaneseCounting)
        {
            sb.Write(@"\levelnfc10"); // Kanji numbering without the digit character (DBNUM1)
        }
        else if (numberingFormat.Val == NumberFormatValues.ChineseLegalSimplified ||
                 numberingFormat.Val == NumberFormatValues.IdeographLegalTraditional ||
                 numberingFormat.Val == NumberFormatValues.JapaneseCounting ||
                 numberingFormat.Val == NumberFormatValues.KoreanCounting)
        {
            sb.Write(@"\levelnfc11"); // Kanji numbering with the digit character (DBNUM2)
        }
        else if (numberingFormat.Val == NumberFormatValues.Aiueo)
        {
            sb.Write(@"\levelnfc12"); // 46 phonetic katakana characters in "aiueo" order (AIUEO) (newer form – “あいうえお。。。” based on phonem matrix) 
        }
        else if (numberingFormat.Val == NumberFormatValues.Iroha)
        {
            sb.Write(@"\levelnfc13"); // 46 phonetic katakana characters in "iroha" order (IROHA) (old form – “いろはにほへとちりぬるお。。。” based on haiku from long ago) 
        }
        else if (numberingFormat.Val == NumberFormatValues.ChineseCountingThousand ||
                numberingFormat.Val == NumberFormatValues.JapaneseLegal ||
                numberingFormat.Val == NumberFormatValues.KoreanLegal ||
                numberingFormat.Val == NumberFormatValues.TaiwaneseCountingThousand)
        {
            sb.Write(@"\levelnfc14"); // Kanji numbering 3 (DBNUM3)
        }
        else if (numberingFormat.Val == NumberFormatValues.KoreanDigital2 ||
                numberingFormat.Val == NumberFormatValues.TaiwaneseDigital)
        {
            sb.Write(@"\levelnfc15"); // Kanji numbering 4 (DBNUM4)
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedCircle)
        {
            sb.Write(@"\levelnfc18"); // Circle numbering (CIRCLENUM) 
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalFullWidth || 
                 numberingFormat.Val == NumberFormatValues.DecimalFullWidth2)
        {
            sb.Write(@"\levelnfc19"); // Double-byte Arabic numbering 
        }
        else if (numberingFormat.Val == NumberFormatValues.Bullet)
        {
            sb.Write(@"\levelnfc23"); // Bullet (no number)
        }
        else if (numberingFormat.Val == NumberFormatValues.Ganada)
        {
            sb.Write(@"\levelnfc24"); // Korean numbering 2 (GANADA) 
        }
        else if (numberingFormat.Val == NumberFormatValues.Chosung)
        {
            sb.Write(@"\levelnfc25"); // Korean numbering 1 (Chosung) 
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedFullstop)
        {
            sb.Write(@"\levelnfc26"); // Chinese numbering 1 (GB1)
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedParen)
        {
            sb.Write(@"\levelnfc27"); // Chinese numbering 2 (GB2)
        }
        else if (numberingFormat.Val == NumberFormatValues.DecimalEnclosedCircleChinese)
        {
            sb.Write(@"\levelnfc28"); // Chinese numbering 3 (GB3)
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographEnclosedCircle)
        {
            sb.Write(@"\levelnfc29"); // Chinese numbering 4 (GB4)
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographTraditional)
        {
            sb.Write(@"\levelnfc30"); // Chinese Zodiac numbering 1 (ZODIAC1) 
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographZodiac)
        {
            sb.Write(@"\levelnfc31"); // Chinese Zodiac numbering 2 (ZODIAC2) 
        }
        else if (numberingFormat.Val == NumberFormatValues.IdeographZodiacTraditional)
        {
            sb.Write(@"\levelnfc32"); // Chinese Zodiac numbering 3 (ZODIAC3) 
        }
        else if (numberingFormat.Val == NumberFormatValues.Hebrew1)
        {
            sb.Write(@"\levelnfc45"); // Hebrew non-standard decimal 
        }
        else if (numberingFormat.Val == NumberFormatValues.ArabicAlpha)
        {
            sb.Write(@"\levelnfc46"); // Arabic Alif Ba Tah
        }
        else if (numberingFormat.Val == NumberFormatValues.Hebrew2)
        {
            sb.Write(@"\levelnfc47"); // Hebrew Biblical standard
        }
        else if (numberingFormat.Val == NumberFormatValues.ArabicAbjad)
        {
            sb.Write(@"\levelnfc48"); // Arabic Abjad style 
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiVowels)
        {
            sb.Write(@"\levelnfc49");
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiConsonants)
        {
            sb.Write(@"\levelnfc50"); 
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiNumbers)
        {
            sb.Write(@"\levelnfc51");
        }
        else if (numberingFormat.Val == NumberFormatValues.HindiCounting)
        {
            sb.Write(@"\levelnfc52"); // Hindi descriptive (cardinals) 
        }
        else if (numberingFormat.Val == NumberFormatValues.ThaiLetters)
        {
            sb.Write(@"\levelnfc53");
        }
        else if (numberingFormat.Val == NumberFormatValues.ThaiNumbers)
        {
            sb.Write(@"\levelnfc54");
        }
        else if (numberingFormat.Val == NumberFormatValues.ThaiCounting)
        {
            sb.Write(@"\levelnfc55"); // Thai descriptive (cardinals) 
        }
        else if (numberingFormat.Val == NumberFormatValues.VietnameseCounting)
        {
            sb.Write(@"\levelnfc56"); // Vietnamese descriptive (cardinals) 
        }
        else if (numberingFormat.Val == NumberFormatValues.NumberInDash)
        {
            sb.Write(@"\levelnfc57"); // Page number format - # -
        }
        else if (numberingFormat.Val == NumberFormatValues.RussianLower)
        {
            sb.Write(@"\levelnfc58"); 
        }
        else if (numberingFormat.Val == NumberFormatValues.RussianUpper)
        {
            sb.Write(@"\levelnfc59"); 
        }
        else if (numberingFormat.Val == NumberFormatValues.Custom)
        {
            if (numberingFormat.Format != null)
            {
                switch (numberingFormat.Format.Value)
                {
                    // These are few standard formats created by Microsoft Word
                    case "01, 02, 03, ...":
                        sb.Write(@"\levelnfc22");
                        break;
                    case "001, 002, 003, ...":
                        sb.Write(@"\levelnfc62");
                        break;
                    case "0001, 0002, 0003, ...":
                        sb.Write(@"\levelnfc63");
                        break;
                    case "00001, 00002, 00003, ...":
                        sb.Write(@"\levelnfc64"); 
                        break;
                }
            }
        }
        else if (numberingFormat.Val == NumberFormatValues.None)
        {
            sb.Write(@"\levelnfc255");
        }
        else
        {
            // Default (decimal numbers)
            sb.Write(@"\levelnfc0");
        }
    }

    private void ProcessLevelText(LevelText levelText, string levelTemplateId, RtfStringWriter sb)
    {
        sb.Write(@"{\leveltext");
        sb.Write(levelTemplateId);
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
            sb.Write($@"\'{rtfParts.Count:X2}");
            foreach (var part in rtfParts)
                sb.Write(part);
        }
        sb.Write(";}");
        
        WriteLevelNumbers(src, sb);
    }

    private void WriteLevelNumbers(string src, RtfStringWriter sb)
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

        sb.Write("{\\levelnumbers");
        foreach (var pos in offsets)
        {
            sb.Write($"\\'{pos:X2}");
        }
        sb.Write(";}");
    }

    private void ProcessListOverrideTable(Numbering numbering, RtfStringWriter sb)
    {
        foreach (var num in numbering.Elements<NumberingInstance>())
        {
            sb.Write(@"{\listoverride"); 

            // Get list id from the AbstractNum element
            if (num.AbstractNumId?.Val != null &&
                numbering.Elements<AbstractNum>().FirstOrDefault(x => x.AbstractNumberId != null && 
                                                                      x.AbstractNumberId == num.AbstractNumId.Val) 
                                                  is AbstractNum abstractNum && 
                abstractNum.Nsid?.Val != null)
            {
                sb.Write(@$"\listid{abstractNum.Nsid.Val.ToLong()}");
            }

            if (num.NumberID != null && num.NumberID.HasValue)
            {
                sb.Write(@$"\ls{num.NumberID.Value}");
            }

            var levelOverrides = num.Elements<LevelOverride>();
            if (levelOverrides == null || !levelOverrides.Any())
            {
                sb.Write(@"\listoverridecount0");
            }
            else
            {
                if (levelOverrides.Count() == 1)
                {
                    sb.Write(@"\listoverridecount1");
                }
                else
                {
                    sb.Write(@"\listoverridecount9");
                }

                foreach (var levelOverride in levelOverrides)
                {
                    sb.Write(@"{\lfolevel");
                    if (levelOverride.StartOverrideNumberingValue?.Val != null)
                    {
                        sb.Write(@$"\listoverridestartat{levelOverride.StartOverrideNumberingValue.Val} ");
                    }
                    /* TODO: if both the start-at and the format are overridden, 
                     * put the \levelstartatN inside the \listlevel contained in the \lfolevel
                     */
                    if (levelOverride.Level is Level level)
                    {
                        sb.Write(@"\listoverrideformat");
                        // TODO: 1, 9 or 0 (not added by Word ?)

                        ProcessLevel(level, sb);
                    }
                    //if (levelOverride.LevelIndex != null)
                    //{
                    // Not supported in RTF, add levels in order instead
                    //}
                    sb.Write('}');
                }
            }
            sb.WriteLine('}');
        }
    }

    internal void ProcessListItem(NumberingProperties numPr, RtfStringWriter sb)
    {
        if (numPr.NumberingLevelReference?.Val != null && numPr.NumberingId?.Val != null)
        {
            // TODO: Calculate list text for RTF readers that don't support automatic numbering
            // or Word 97-2007 (or newer) lists. For now, just use a generic bullet.
            fonts.TryAddAndGetIndex("Arial", out int fontIndex);
            sb.Write($@"{{\listtext \f{fontIndex}\bullet\tab}}");

            sb.Write($@"\ls{numPr.NumberingId.Val}\ilvl{numPr.NumberingLevelReference.Val}");
        }
    }
}
