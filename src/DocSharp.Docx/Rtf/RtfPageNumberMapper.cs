using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Rtf;

internal static class RtfPageNumberMapper
{
    internal static NumberFormatValues? GetPageNumberFormat(string format)
    {
        switch(format)
        {
            case "pgndec":
                return NumberFormatValues.Decimal;
            case "pgnlcltr":
                return NumberFormatValues.LowerLetter;
            case "pgnucltr":
                return NumberFormatValues.UpperLetter;
            case "pgnlcrm":
                return NumberFormatValues.LowerRoman;
            case "pgnucrm":
                return NumberFormatValues.UpperRoman;
            case "pgncnum": // Circle numbering (CIRCLENUM)
                return NumberFormatValues.DecimalEnclosedCircle;

            // For the following I am not sure about the mapping: 
            case "pgnbidia": // TODO: page-number format is Alif Ba Tah if language is Arabic and Non-standard Decimal if language is Hebrew
                return NumberFormatValues.ArabicAlpha;
            case "pgnbidib": // TODO: page-number format is Alif Ba Tah if language is Arabic and Non-standard Decimal if language is Hebrew
                return NumberFormatValues.ArabicAbjad;
            case "pgnchosung": // Korean numbering 1 (CHOSUNG)
                return NumberFormatValues.Chosung;
            case "pgndbnum": // Kanji numbering without the digit character
                return NumberFormatValues.ChineseCounting;
            case "pgndbnumd": // Kanji numbering with the digit character
                return NumberFormatValues.JapaneseCounting;
            case "pgndbnumt": // Kanji numbering 3 (DBNUM3)
                return NumberFormatValues.ChineseCountingThousand;
            case "pgndbnumk": // Kanji numbering 4 (DBNUM4)
                return NumberFormatValues.KoreanDigital2;
            case "pgndecd": // Double-byte decimal numbering
                return NumberFormatValues.DecimalFullWidth;
            case "pgnganada": // Korean numbering 2 (GANADA)
                return NumberFormatValues.Ganada;
            case "pgngbnum": // Chinese numbering 1 (GB1)
                return NumberFormatValues.DecimalEnclosedFullstop;
            case "pgngbnumd": // Chinese numbering 2 (GB2)
                return NumberFormatValues.DecimalEnclosedParen;
            case "pgngbnuml": // Chinese numbering 3 (GB3)
                return NumberFormatValues.DecimalEnclosedCircleChinese;
            case "pgngbnumk": // Chinese numbering 4 (GB4)
                return NumberFormatValues.IdeographEnclosedCircle;
            case "pgnzodiac": // Chinese Zodiac numbering 1 (ZODIAC1)
                return NumberFormatValues.IdeographTraditional;
            case "pgnzodiacd": // Chinese Zodiac numbering 2 (ZODIAC2)
                return NumberFormatValues.IdeographZodiac;
            case "pgnzodiacl": // Chinese Zodiac numbering 3 (ZODIAC3)
                return NumberFormatValues.IdeographZodiacTraditional;
            case "pgnhindia": // Hindi vowel numeric format
                return NumberFormatValues.HindiVowels;
            case "pgnhindib": // Hindi consonants
                return NumberFormatValues.HindiConsonants;
            case "pgnhindic": // Hindi digits
                return NumberFormatValues.HindiNumbers;
            case "pgnhindid": // Hindi descriptive (cardinal) text
                return NumberFormatValues.HindiCounting;
            case "pgnthaia": // Thai letters
                return NumberFormatValues.ThaiLetters;
            case "pgnthaib": // Thai digits
                return NumberFormatValues.ThaiNumbers;
            case "pgnthaic": // Thai descriptive
                return NumberFormatValues.ThaiCounting;
            case "pgnvieta": // Vietnamese  descriptive
                return NumberFormatValues.VietnameseCounting;
            case "pgnid": // Page number in dashes (Korean)
                return NumberFormatValues.NumberInDash;

            default:
                return null;
        }
    }
}
