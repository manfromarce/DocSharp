using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal void ProcessListItem(NumberingProperties numPr, StringBuilder sb)
    {
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;
            var num = numberingPart.Elements<NumberingInstance>()
                                   .FirstOrDefault(x => x.NumberID == numPr.NumberingId.Val);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                  .FirstOrDefault(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null && 
                                                                               x.LevelIndex == levelIndex);
                if (level != null && 
                    level.NumberingFormat?.Val is EnumValue<NumberFormatValues> listType && 
                    listType != NumberFormatValues.None &&
                    level.LevelText?.Val?.Value is string levelText)
                {
                    // This part is not used by most RTF readers.
                    //sb.Append(@"{\pntext\f0\'95\tab}");
                    ////

                    sb.Append(@"{\*\pn");

                    bool isBulleted = listType == NumberFormatValues.Bullet;
                    if (isBulleted)
                    {
                        sb.Append(@"\pnlvlblt");
                    }
                    else
                    {
                        sb.Append(@"\pnlvlbody");
                    }

                    var font = level.NumberingSymbolRunProperties?.RunFonts?.Ascii?.Value;
                    if (font != null)
                    {
                        fonts.TryAddAndGetIndex(font, out int fontIndex);
                        sb.Append(@"\pnf" + fontIndex);
                    }

                    if (level.PreviousParagraphProperties?.Indentation != null &&
                        int.TryParse(level.PreviousParagraphProperties.Indentation.Hanging?.Value, out int hanging))
                    {
                        sb.Append(@"\pnindent" + hanging);
                        //sb.Append(@"\pnsp");
                    }

                    var start = level.StartNumberingValue?.Val ?? 1;
                    sb.Append(@"\pnstart" + start);

                    if (isBulleted && levelText.Length > 0)
                    {
                        // Write bullet char code
                        sb.Append(@"{\pntxtb " + "\\u" + (int)levelText[0] + "?" + "}");
                    }
                    else if (listType == NumberFormatValues.Aiueo)
                    {
                        sb.Append(@"\pnaiu"); // 46 phonetic katakana characters in "aiueo" order (AIUEO)
                    }
                    else if (listType == NumberFormatValues.AiueoFullWidth)
                    {
                        sb.Append(@"\pnaiud"); // 46 phonetic double-byte katakana characters (AIUEO DBCHAR)
                    }
                    else if (listType == NumberFormatValues.ArabicAbjad || 
                             listType == NumberFormatValues.Hebrew2)
                    {
                        sb.Append(@"\pnbidib"); // Alif Ba Tah if language is Arabic and Non-standard Decimal if language is Hebrew
                    }
                    else if (listType == NumberFormatValues.ArabicAlpha || 
                             listType == NumberFormatValues.Hebrew1)
                    {
                        sb.Append(@"\pnbidia"); // Abjad Jawaz if language is Arabic and Biblical Standard if language is Hebrew
                    }
                    else if (listType == NumberFormatValues.ChineseCounting ||
                             listType == NumberFormatValues.IdeographDigital ||
                             listType == NumberFormatValues.KoreanDigital ||
                             listType == NumberFormatValues.TaiwaneseCounting)
                    {
                        sb.Append(@"\pndbnum"); // Kanji numbering without the digit character (DBNUM1)
                    }
                    else if (listType == NumberFormatValues.ChineseLegalSimplified ||
                             listType == NumberFormatValues.IdeographLegalTraditional ||
                             listType == NumberFormatValues.JapaneseCounting ||
                             listType == NumberFormatValues.KoreanCounting)
                    {
                        sb.Append(@"\pndbnumd"); // Kanji numbering with the digit character (DBNUM2)
                    }
                    else if (listType == NumberFormatValues.ChineseCountingThousand ||
                             listType == NumberFormatValues.JapaneseLegal ||
                             listType == NumberFormatValues.KoreanLegal ||
                             listType == NumberFormatValues.TaiwaneseCountingThousand)
                    {
                        sb.Append(@"\pndbnumt"); // Kanji numbering 3 (DBNUM3), alias for \pndbnuml
                    }
                    else if (listType == NumberFormatValues.KoreanDigital2 ||
                            listType == NumberFormatValues.TaiwaneseDigital)
                    {
                        sb.Append(@"\pndbnumk"); // Kanji numbering 4 (DBNUM4)
                    }
                    else if (listType == NumberFormatValues.CardinalText)
                    {
                        sb.Append(@"\pncard"); // One, Two, Three
                    }
                    else if (listType == NumberFormatValues.Chosung)
                    {
                        sb.Append(@"\pnchosung"); // Korean numbering 1 (CHOSUNG)
                    }                    
                    else if (listType == NumberFormatValues.DecimalEnclosedCircle)
                    {
                        sb.Append(@"\pncnum"); // 20 numbered list in circle (CIRCLENUM)
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedFullstop)
                    {
                        sb.Append(@"\pngbnum"); // Chinese numbering 1 (GB1)
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedParen)
                    {
                        sb.Append(@"\pngbnumd"); // Chinese numbering 2 (GB2)
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedCircleChinese)
                    {
                        sb.Append(@"\pngbnuml"); // Chinese numbering 3 (GB3)
                    }
                    else if (listType == NumberFormatValues.IdeographEnclosedCircle)
                    {
                        sb.Append(@"\pngbnumk"); // Chinese numbering 4 (GB4)
                    }
                    else if (listType == NumberFormatValues.DecimalFullWidth ||
                             listType == NumberFormatValues.DecimalFullWidth2)
                    {
                        sb.Append(@"\pndecd"); // Double-byte decimal numbering (Arabic DBCHAR)
                    }
                    else if (listType == NumberFormatValues.Ganada)
                    {
                        sb.Append(@"\pnganada"); // Korean numbering 2 (GANADA)
                    }
                    else if (listType == NumberFormatValues.Iroha)
                    {
                        sb.Append(@"\pniroha"); // 46 phonetic katakana characters in "iroha" order (IROHA)
                    }
                    else if (listType == NumberFormatValues.IrohaFullWidth)
                    {
                        sb.Append(@"\pnirohad"); // 46 phonetic double-byte katakana characters (IROHA DBCHAR)
                    }
                    else if (listType == NumberFormatValues.LowerLetter)
                    {
                        sb.Append(@"\pnlcltr"); // a, b, c
                    }
                    else if (listType == NumberFormatValues.LowerRoman)
                    {
                        sb.Append(@"\pnlcrm"); // i, ii, iii
                    }
                    else if (listType == NumberFormatValues.Ordinal)
                    {
                        sb.Append(@"\pnord"); // 1st, 2nd, 3rd
                    }
                    else if (listType == NumberFormatValues.OrdinalText)
                    {
                        sb.Append(@"\pnordt"); // First, Second, Third
                    }
                    else if (listType == NumberFormatValues.UpperLetter)
                    {
                        sb.Append(@"\pnucltr"); // A, B, C
                    }
                    else if (listType == NumberFormatValues.UpperRoman)
                    {
                        sb.Append(@"\pnucrm"); // I, II, III
                    }
                    else if (listType == NumberFormatValues.IdeographTraditional)
                    {
                        sb.Append(@"\pnzodiac"); //Chinese Zodiac numbering 1 (ZODIAC1)
                    }
                    else if (listType == NumberFormatValues.IdeographZodiac)
                    {
                        sb.Append(@"\pnzodiacd"); //Chinese Zodiac numbering 2 (ZODIAC2)
                    }
                    else if (listType == NumberFormatValues.IdeographZodiacTraditional)
                    {
                        sb.Append(@"\pnzodiacl"); //Chinese Zodiac numbering 3 (ZODIAC3)
                    }
                    else
                    {
                        sb.Append(@"\pndec"); // Decimal numbering (1, 2, 3)
                    }

                    if (!isBulleted)
                    {
                        if (levelText.Contains("%"))
                        {
                            var formatParts = levelText.Split('%');
                            string before = formatParts[0];
                            string after = formatParts[1].TrimStart(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']);
                            if (before != string.Empty)
                            {
                                sb.Append(@"{\pntxtb " + before + "}");
                            }
                            if (after != string.Empty)
                            {
                                sb.Append(@"{\pntxta " + after + "}");
                            }
                        }
                    }

                    sb.Append('}');

                    if (level.PreviousParagraphProperties?.Indentation != null &&
                        int.TryParse(level.PreviousParagraphProperties.Indentation.Left?.Value, out int left))
                    {
                        sb.Append(@"\fi" + left);
                    }
                }
            }
        }
    }
}
