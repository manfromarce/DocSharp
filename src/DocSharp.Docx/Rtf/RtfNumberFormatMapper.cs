using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Rtf;

internal static class RtfNumberFormatMapper
{
    internal static NumberFormatValues? GetNumberFormat(string format)
    {
        switch(format)
        {
            case "pgndec": // Page numbers
            case "pndec": // Legacy lists (not using list table)
            case "levelnfc0": // Lists (level in the list table)
            case "ftnnar": // Foonotes (document level)
            case "aftnnar": // Endnotes (document level)
            case "sftnnar": // Foonotes (section level)
            case "saftnnar": // Endnotes (section level)
            // Note that not all formats are available for all element types (page numbers, lists, footnotes/endnotes) in RTF.
                return NumberFormatValues.Decimal;
            case "pgnucrm":
            case "pnucrm":
            case "levelnfc1":
            case "ftnnruc":
            case "aftnnruc":
            case "sftnnruc":
            case "saftnnruc":
                return NumberFormatValues.UpperRoman;
            case "pgnlcrm":
            case "pnlcrm":
            case "levelnfc2":
            case "ftnnrlc":
            case "aftnnrlc":
            case "sftnnrlc":
            case "saftnnrlc":
                return NumberFormatValues.LowerRoman;
            case "pgnucltr":
            case "pnucltr":
            case "levelnfc3":
            case "ftnnauc": 
            case "aftnnauc":
            case "sftnnauc":
            case "saftnnauc":
                return NumberFormatValues.UpperLetter;
            case "pgnlcltr":
            case "pnlcltr": 
            case "levelnfc4":
            case "ftnnalc":
            case "aftnnalc":
            case "sftnnalc":
            case "saftnnalc":
                return NumberFormatValues.LowerLetter;
            case "pnord": // Ordinal numbering (1st, 2nd, 3rd)
            case "levelnfc5":
                return NumberFormatValues.Ordinal;
            case "pncard": // Cardinal numbering (One, Two, Three)
            case "levelnfc6":
                return NumberFormatValues.CardinalText;
            case "pnordt": // Ordinal text numbering (First, Second, Third)
            case "levelnfc7":
                return NumberFormatValues.OrdinalText;
            case "levelnfc8":
                return NumberFormatValues.Hex;
            case "ftnnchi": // Chicago Manual of Style (*, †, ‡, §)
            case "levelnfc9":
            case "levelnfc70":
            case "aftnnchi":
            case "sftnnchi":
            case "saftnnchi":
                return NumberFormatValues.Chicago;
            case "pgndbnum": // Kanji numbering without the digit character (DBNUM1)
            case "pndbnum":
            case "levelnfc10":
            case "ftnndbnum":
            case "aftnndbnum":
            case "sftnndbnum":
            case "saftnndbnum":
                return NumberFormatValues.IdeographDigital;
            case "pgndbnumd": // Kanji numbering with the digit character (DBNUM2)
            case "pndbnumd":
            case "levelnfc11":
            case "ftnndbnumd":
            case "aftnndbnumd":
            case "sftnndbnumd":
            case "saftnndbnumd":
                return NumberFormatValues.JapaneseCounting;
            case "pnaiu": // 46 phonetic katakana characters in "aiueo" order
            case "pnaiueo":
            case "levelnfc12":
                return NumberFormatValues.Aiueo;
            case "pniroha": // 46 phonetic katakana characters in "iroha" order
            case "levelnfc13":
                return NumberFormatValues.Iroha;
            case "pgndecd": // Double-byte decimal numbering
            case "pndecd":
            case "levelnfc14":
            case "ftnndbar":
            case "aftnndbar":
            case "sftnndbar":
            case "saftnndbar":
                return NumberFormatValues.DecimalFullWidth;
            case "levelnfc15":
                return NumberFormatValues.DecimalHalfWidth;
            case "pgndbnumt": // Kanji numbering 3 (DBNUM3)
            case "pndbnuml":
            case "pndbnumt": // alias for \pndbnuml
            case "levelnfc16":
            case "ftnngbnumt":
            case "aftnndbnumt":
            case "sftnndbnumt":
            case "saftnndbnumt":
                return NumberFormatValues.JapaneseLegal;
            case "pgndbnumk": // Kanji numbering 4 (DBNUM4)
            case "pndbnumk":
            case "levelnfc17":
            case "ftnndbnumk":
            case "aftnndbnumk":
            case "sftnndbnumk":
            case "saftnndbnumk":
                return NumberFormatValues.JapaneseDigitalTenThousand;
            case "pgncnum": // Circle numbering (CIRCLENUM)
            case "pncnum": 
            case "levelnfc18":
            case "ftnncnum":
            case "aftnncnum":
            case "sftnncnum":
            case "saftnncnum":
                return NumberFormatValues.DecimalEnclosedCircle;
            case "levelnfc19":
                return NumberFormatValues.DecimalFullWidth2;
            case "pnaiud": // 46 phonetic double-byte katakana characters (AIUEO DBCHAR)
            case "pnaiueod":
            case "levelnfc20":
                return NumberFormatValues.AiueoFullWidth;
            case "pnirohad": // 46 phonetic double-byte katakana characters (IROHA DBCHAR)
            case "levelnfc21":
                return NumberFormatValues.IrohaFullWidth;
            case "levelnfc22": // Arabic with leading zero (01, 02, 03, ..., 10, 11)
                return NumberFormatValues.DecimalZero;
            case "levelnfc23": // Bullet (no number)
                return NumberFormatValues.Bullet;
            case "pgnganada": // Korean numbering 2 (GANADA)
            case "pnganada":
            case "levelnfc24":
            case "ftnnganada":
            case "aftnnganada":
            case "sftnnganada":
            case "saftnnganada":
                return NumberFormatValues.Ganada;
            case "pgnchosung": // Korean numbering 1 (CHOSUNG)
            case "pnchosung":
            case "levelnfc25":
            case "ftnnchosung":
            case "aftnnchosung":
            case "sftnnchosung":
            case "saftnnchosung":
                return NumberFormatValues.Chosung;
            case "pgngbnum": // Chinese numbering 1 (GB1)
            case "pngbnum":
            case "levelnfc26":
            case "ftnngbnum":
            case "aftnngbnum":
            case "sftnngbnum":
            case "saftnngbnum":
                return NumberFormatValues.DecimalEnclosedFullstop;
            case "pgngbnumd": // Chinese numbering 2 (GB2)
            case "pngbnumd":
            case "levelnfc27":
            case "ftnngbnumd":
            case "aftnngbnumd":
            case "sftnngbnumd":
            case "saftnngbnumd":
                return NumberFormatValues.DecimalEnclosedParen;
            case "pgngbnuml": // Chinese numbering 3 (GB3)
            case "pngbnuml":
            case "levelnfc28":
            case "ftnngbnuml":
            case "aftnngbnuml":
            case "sftnngbnuml":
            case "saftnngbnuml":
                return NumberFormatValues.DecimalEnclosedCircleChinese;
            case "pgngbnumk": // Chinese numbering 4 (GB4)
            case "pngbnumk":
            case "levelnfc29":
            case "ftnngbnumk":
            case "aftnngbnumk":
            case "sftnngbnumk":
            case "saftnngbnumk":
                return NumberFormatValues.IdeographEnclosedCircle;
            case "pgnzodiac": // Chinese Zodiac numbering 1 (ZODIAC1)
            case "pnzodiac":
            case "levelnfc30":
            case "ftnnzodiac":
            case "aftnnzodiac":
            case "sftnnzodiac":
            case "saftnnzodiac":
                return NumberFormatValues.IdeographTraditional;
            case "pgnzodiacd": // Chinese Zodiac numbering 2 (ZODIAC2)
            case "pnzodiacd":
            case "levelnfc31":
            case "ftnnzodiacd":
            case "aftnnzodiacd":
            case "sftnnzodiacd":
            case "saftnnzodiacd":
                return NumberFormatValues.IdeographZodiac;
            case "pgnzodiacl": // Chinese Zodiac numbering 3 (ZODIAC3)
            case "pnzodiacl":
            case "levelnfc32":
            case "ftnnzodiacl":
            case "aftnnzodiacl":
            case "sftnnzodiacl":
            case "saftnnzodiacl":
                return NumberFormatValues.IdeographZodiacTraditional;
            case "levelnfc33":
                return NumberFormatValues.TaiwaneseCounting;
            case "levelnfc34":
                return NumberFormatValues.IdeographLegalTraditional;
            case "levelnfc35":
                return NumberFormatValues.TaiwaneseCountingThousand;
            case "levelnfc36":
                return NumberFormatValues.TaiwaneseDigital;
  
            case "levelnfc37":
            case "levelnfc40":
                return NumberFormatValues.ChineseCounting;
            case "levelnfc38":
                return NumberFormatValues.ChineseLegalSimplified;
            case "levelnfc39":
                return NumberFormatValues.ChineseCountingThousand;

            case "levelnfc41":
                return NumberFormatValues.KoreanDigital;
            case "levelnfc42":
                return NumberFormatValues.KoreanCounting;
            case "levelnfc43":
                return NumberFormatValues.KoreanLegal;
            case "levelnfc44":
                return NumberFormatValues.KoreanDigital2;
            case "levelnfc45":
                return NumberFormatValues.Hebrew1;
            case "levelnfc46":
                return NumberFormatValues.ArabicAlpha;
            case "levelnfc47":
                return NumberFormatValues.Hebrew2;
            case "levelnfc48":
                return NumberFormatValues.ArabicAbjad;
            case "pgnhindia": // Hindi vowel numeric format
            case "levelnfc49":
                return NumberFormatValues.HindiVowels;
            case "pgnhindib": // Hindi consonants
            case "levelnfc50":
                return NumberFormatValues.HindiConsonants;
            case "pgnhindic": // Hindi digits
            case "levelnfc51":
                return NumberFormatValues.HindiNumbers;
            case "pgnhindid": // Hindi descriptive (cardinal) text
            case "levelnfc52":
                return NumberFormatValues.HindiCounting;
            case "pgnthaia": // Thai letters
            case "levelnfc53":
                return NumberFormatValues.ThaiLetters;
            case "pgnthaib": // Thai digits
                return NumberFormatValues.ThaiNumbers;
            case "pgnthaic": // Thai descriptive
            case "levelnfc55":
                return NumberFormatValues.ThaiCounting;
            case "pgnvieta": // Vietnamese  descriptive
            case "levelnfc56":
                return NumberFormatValues.VietnameseCounting;
            case "pgnid": // Page number in dashes (Korean)
            case "levelnfc57":
                return NumberFormatValues.NumberInDash;
            case "levelnfc58": // Lowercase Russian alphabet
                return NumberFormatValues.RussianLower;
            case "levelnfc59": // Uppercase Russian alphabet
                return NumberFormatValues.RussianUpper;

            case "levelnfc60": // Lowercase Greek numerals (alphabet based)
            case "levelnfc61": // Uppercase Greek numerals (alphabet based)
            case "levelnfc62": // 2 leading zeros: 001, 002, ..., 100, ...
            case "levelnfc63": // 3 leading zeros: 0001, 0002, ..., 1000, ...
            case "levelnfc64": // 4 leading zeros: 00001, 00002, ..., 10000, ...
            case "levelnfc65": // Lowercase Turkish alphabet
            case "levelnfc66": // Uppercase Turkish alphabet
            case "levelnfc67": // Lowercase Bulgarian alphabet
            case "levelnfc68": // Uppercase Bulgarian alphabet
                // These are not available in Open XML
                return NumberFormatValues.Decimal;

            case "levelnfc255": // No number
                return NumberFormatValues.None;

            case "pgnbidia": 
            case "pnbidia": 
                // TODO: page-number format is Alif Ba Tah if language is Arabic and Non-standard Decimal if language is Hebrew
                return NumberFormatValues.ArabicAlpha;
            case "pgnbidib": 
            case "pnbidib": 
                // TODO: page-number format is Alif Ba Tah if language is Arabic and Non-standard Decimal if language is Hebrew
                return NumberFormatValues.ArabicAbjad;            

            default:
                if (format.StartsWith("levelnfc"))
                    // For values that don't have an equivalent in Open XML, use decimal numbers
                    return NumberFormatValues.Decimal;
                else 
                    // For unrecognized control words, return null (will be handled in the caller)
                    return null;
        }
    }
}
