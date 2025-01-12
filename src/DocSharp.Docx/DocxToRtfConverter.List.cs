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
                    }

                    //sb.Append(@"\pnsp");

                    var start = level.StartNumberingValue?.Val ?? 1;
                    sb.Append(@"\pnstart" + start);

                    if (isBulleted && levelText.Length > 0)
                    {
                        // Write bullet char code
                        sb.Append(@"{\pntxtb " + "\\u" + (int)levelText[0] + "?" + "}");
                    }
                    else if (listType == NumberFormatValues.Aiueo)
                    {
                        sb.Append(@"\pnaiu");
                    }
                    else if (listType == NumberFormatValues.AiueoFullWidth)
                    {
                        sb.Append(@"\pnaiud");
                    }
                    else if (listType == NumberFormatValues.ArabicAbjad)
                    {
                    }
                    else if (listType == NumberFormatValues.ArabicAlpha)
                    {
                    }
                    else if (listType == NumberFormatValues.BahtText)
                    {
                    }
                    else if (listType == NumberFormatValues.CardinalText)
                    {
                        sb.Append(@"\pncard");
                    }
                    else if (listType == NumberFormatValues.Chicago)
                    {
                    }
                    else if (listType == NumberFormatValues.ChineseCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.ChineseCountingThousand)
                    {
                    }
                    else if (listType == NumberFormatValues.ChineseLegalSimplified)
                    {
                    }
                    else if (listType == NumberFormatValues.Chosung)
                    {
                        sb.Append(@"\pnchosung");
                    }
                    else if (listType == NumberFormatValues.Decimal)
                    {
                        sb.Append(@"\pndec");
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedCircle)
                    {
                        sb.Append(@"\pncnum");
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedCircleChinese)
                    {
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedFullstop)
                    {
                    }
                    else if (listType == NumberFormatValues.DecimalEnclosedParen)
                    {
                    }
                    else if (listType == NumberFormatValues.DecimalFullWidth)
                    {
                        sb.Append(@"\pndecd");
                    }
                    else if (listType == NumberFormatValues.DecimalFullWidth2)
                    {
                        sb.Append(@"\pndecd");
                    }
                    else if (listType == NumberFormatValues.DecimalHalfWidth)
                    {
                    }
                    else if (listType == NumberFormatValues.DecimalZero)
                    {
                    }
                    else if (listType == NumberFormatValues.DollarText)
                    {
                    }
                    else if (listType == NumberFormatValues.Ganada)
                    {
                        sb.Append(@"\pnganada");
                    }
                    else if (listType == NumberFormatValues.Hebrew1)
                    {
                    }
                    else if (listType == NumberFormatValues.Hebrew2)
                    {
                    }
                    else if (listType == NumberFormatValues.Hex)
                    {
                    }
                    else if (listType == NumberFormatValues.HindiConsonants)
                    {
                    }
                    else if (listType == NumberFormatValues.HindiCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.HindiNumbers)
                    {
                    }
                    else if (listType == NumberFormatValues.HindiVowels)
                    {
                    }
                    else if (listType == NumberFormatValues.IdeographDigital)
                    {
                    }
                    else if (listType == NumberFormatValues.IdeographEnclosedCircle)
                    {
                    }
                    else if (listType == NumberFormatValues.IdeographLegalTraditional)
                    {
                    }
                    else if (listType == NumberFormatValues.IdeographTraditional)
                    {
                    }
                    else if (listType == NumberFormatValues.IdeographZodiac)
                    {
                    }
                    else if (listType == NumberFormatValues.IdeographZodiacTraditional)
                    {
                    }
                    else if (listType == NumberFormatValues.Iroha)
                    {
                        sb.Append(@"\pniroha");
                    }
                    else if (listType == NumberFormatValues.IrohaFullWidth)
                    {
                        sb.Append(@"\pnirohad");
                    }
                    else if (listType == NumberFormatValues.JapaneseCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.JapaneseDigitalTenThousand)
                    {
                    }
                    else if (listType == NumberFormatValues.JapaneseLegal)
                    {
                        sb.Append(@"\pngbnumd"); // ?
                    }
                    else if (listType == NumberFormatValues.KoreanCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.KoreanDigital)
                    {
                    }
                    else if (listType == NumberFormatValues.KoreanDigital2)
                    {
                    }
                    else if (listType == NumberFormatValues.KoreanLegal)
                    {
                    }
                    else if (listType == NumberFormatValues.LowerLetter)
                    {
                        sb.Append(@"\pnlcltr");
                    }
                    else if (listType == NumberFormatValues.LowerRoman)
                    {
                        sb.Append(@"\pnlcrm");
                    }
                    else if (listType == NumberFormatValues.NumberInDash)
                    {
                    }
                    else if (listType == NumberFormatValues.Ordinal)
                    {
                        sb.Append(@"\pnord");
                    }
                    else if (listType == NumberFormatValues.OrdinalText)
                    {
                        sb.Append(@"\pnordt");
                    }
                    else if (listType == NumberFormatValues.RussianLower)
                    {
                    }
                    else if (listType == NumberFormatValues.RussianUpper)
                    {
                    }
                    else if (listType == NumberFormatValues.TaiwaneseCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.TaiwaneseCountingThousand)
                    {
                    }
                    else if (listType == NumberFormatValues.TaiwaneseDigital)
                    {
                    }
                    else if (listType == NumberFormatValues.ThaiCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.ThaiLetters)
                    {
                    }
                    else if (listType == NumberFormatValues.ThaiNumbers)
                    {
                    }
                    else if (listType == NumberFormatValues.UpperLetter)
                    {
                        sb.Append(@"\pnucltr");
                    }
                    else if (listType == NumberFormatValues.UpperRoman)
                    {
                        sb.Append(@"\pnucrm");
                    }
                    else if (listType == NumberFormatValues.VietnameseCounting)
                    {
                    }
                    else if (listType == NumberFormatValues.Custom)
                    {
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
