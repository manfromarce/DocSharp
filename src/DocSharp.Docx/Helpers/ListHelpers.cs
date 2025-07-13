using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class ListHelpers
{
    public static Level? GetListLevel(NumberingProperties? numPr)
    {
        if (numPr == null)
        {
            return null;
        }

        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int numberingId = numPr.NumberingId.Val;
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;

            var num = numberingPart.Elements<NumberingInstance>()
                                .FirstOrDefault(x => x.NumberID != null &&
                                                     x.NumberID == numberingId);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                .FirstOrDefault(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                            x.LevelIndex == levelIndex);
                var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                                    x.LevelIndex == levelIndex);

                // Use LevelOverride if present
                return levelOverride?.Level ?? level;
            }
        }
        return null;
    }

    public static string GetNumberString(string? levelText, EnumValue<NumberFormatValues> listType, int numberingId, int levelIndex, Dictionary<(int NumberingId, int LevelIndex), int> _listLevelCounters, CultureInfo? culture = null)
    {
        if (listType == NumberFormatValues.Bullet)
        {
            return "•";
        }

        if (levelText != null)
        {
            string formattedText = levelText;
            // TODO: use the document language
            culture ??= CultureInfo.InvariantCulture;
            // culture ??= CultureInfo.CurrentCulture;
            foreach (var kvp in _listLevelCounters.Where(k => k.Key.NumberingId == numberingId))
            {
                var placeholder = kvp.Key.LevelIndex + 1;
                string value;
                if (listType == NumberFormatValues.LowerLetter)
                {
                    value = ListHelpers.NumberToLetter(kvp.Value, false);
                }
                else if (listType == NumberFormatValues.UpperLetter)
                {
                    value = ListHelpers.NumberToLetter(kvp.Value, true);
                }
                else if (listType == NumberFormatValues.LowerRoman)
                {
                    value = ListHelpers.NumberToRomanLetter(kvp.Value, false);
                }
                else if (listType == NumberFormatValues.UpperRoman)
                {
                    value = ListHelpers.NumberToRomanLetter(kvp.Value, true);
                }
                else if (listType == NumberFormatValues.DecimalEnclosedCircle || 
                         listType == NumberFormatValues.DecimalEnclosedCircleChinese)
                {
                    value = ListHelpers.NumberToCircledNumber(kvp.Value, true);
                }
                else if (listType == NumberFormatValues.Chicago)
                {
                    value = ListHelpers.NumberToChicago(kvp.Value);
                }
                else if (listType == NumberFormatValues.Hex)
                {
                    value = kvp.Value.ToString("X", CultureInfo.InvariantCulture); // Use invariant culture for hex format
                }
                else if (listType == NumberFormatValues.NumberInDash)
                {
                    value = $"- {kvp.Value.ToString(culture ?? CultureInfo.InvariantCulture)} -";
                }
                else if (listType == NumberFormatValues.DecimalEnclosedFullstop)
                {
                    value = $"{kvp.Value.ToString(culture ?? CultureInfo.InvariantCulture)}.";
                }
                else if (listType == NumberFormatValues.DecimalEnclosedParen)
                {
                    value = $"({kvp.Value.ToString(culture ?? CultureInfo.InvariantCulture)})";
                }
                else if (listType == NumberFormatValues.DecimalZero)
                {
                    value = kvp.Value.ToString("00");
                }
                // else if (listType == NumberFormatValues.Ordinal)
                // {
                // }
                // else if (listType == NumberFormatValues.OrdinalText)
                // {
                // }
                // else if (listType == NumberFormatValues.CardinalText)
                // {
                // }
                else if (listType == NumberFormatValues.Bullet)
                {
                    value = "•";
                }
                else if (listType == NumberFormatValues.None)
                {
                    value = string.Empty;
                }
                else
                {
                    // Regular number
                    value = kvp.Value.ToString(culture);
                }
                formattedText = formattedText.Replace($"%{placeholder}", value);
            }
            return formattedText;
        }

        return _listLevelCounters[(numberingId, levelIndex)].ToString();
    }

    public static string NumberToLetter(int number, bool uppercase)
    {
        if (number < 1)
            number = 1;

        const int alphabetLength = 26;
        number--; // Convert to zero-based index
        char letter = (char)('a' + (number % alphabetLength));
        string result = new string(letter, 1);
        if (number >= alphabetLength)
        {
            result = NumberToLetter(number / alphabetLength, uppercase) + result;
        }
        return uppercase ? result.ToUpperInvariant() : result;
    }

    public static string NumberToChicago(int number)
    {
        if (number < 0)
            number = 1;

        string[] symbols = { "*", "†", "‡", "§" };
        int index = (number - 1) % symbols.Length;
        int repetitions = (number - 1) / symbols.Length + 1;
        return new string(symbols[index][0], repetitions);
    }


    public static string NumberToCircledNumber(int number, bool uppercase)
    {
        if (number < 0)
            number = 1;
        if (number > 20)
            number = 20;

        return number switch
        {
            0 => "⓪",
            1 => "①",
            2 => "②",
            3 => "③",
            4 => "④",
            5 => "⑤",
            6 => "⑥",
            7 => "⑦",
            8 => "⑧",
            9 => "⑨",
            10 => "⑩",
            11 => "⑪",
            12 => "⑫",
            13 => "⑬",
            14 => "⑭",
            15 => "⑮",
            16 => "⑯",
            17 => "⑰",
            18 => "⑱",
            19 => "⑲",
            20 => "⑳",
            _ => number.ToString(),
        };
    }

    public static string NumberToRomanLetter(long number, bool uppercase)
    {
        if (number < 1)
            number = 1;

        if (number > 3999)
            number = 3999;

        string[] romanSymbols = { "m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i" };
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string result = "";

        for (int i = 0; i < romanSymbols.Length; i++)
        {
            while (number >= values[i])
            {
                result += romanSymbols[i];
                number -= values[i];
            }
        }
        return uppercase ? result.ToUpperInvariant() : result;
    }

    public static AbstractNum AddOrderedListAbstractNumbering(this WordprocessingDocument document)
    {
        var numbering = document.GetOrCreateNumbering();

        var abstractNumId = numbering.Elements<AbstractNum>().Count() + 1;

        var abstractNum = new AbstractNum(
            new Level(
                new NumberingFormat() { Val = NumberFormatValues.Decimal },
                new LevelText() { Val = "%1." }
            )
            { LevelIndex = 0, StartNumberingValue = new StartNumberingValue() { Val = 1 } }
        )
        {
            AbstractNumberId = abstractNumId,
            MultiLevelType = new MultiLevelType { Val = MultiLevelValues.SingleLevel }
        };

        numbering.AddAbstractNumbering(abstractNum);

        return abstractNum;
    }

    public static AbstractNum AddBulletListAbstractNumbering(this WordprocessingDocument document)
    {
        var numbering = document.GetOrCreateNumbering();

        var abstractNumberId = numbering.Elements<AbstractNum>().Count() + 1;

        var abstractNum = new AbstractNum(
            new Level(
                new NumberingFormat() { Val = NumberFormatValues.Bullet },
                new LevelText() { Val = "·" }
            )
            { LevelIndex = 0 }
        )
        { AbstractNumberId = abstractNumberId };

        numbering.AddAbstractNumbering(abstractNum);
        return abstractNum;
    }

    public static NumberingInstance AddOrderedListNumbering(this WordprocessingDocument document, int abstractNumId, int? startFrom = null)
    {
        var numbering = document.GetOrCreateNumbering();
        var numId = numbering.Elements<NumberingInstance>().Count() + 1;
        var numberingInstance = new NumberingInstance(
            new AbstractNumId() { Val = abstractNumId }
        )
        { NumberID = numId };
        numbering.AddNumberingInstance(numberingInstance);

        if (startFrom != null)
        {
            var levelOverride = new LevelOverride
            {
                LevelIndex = 0,
                StartOverrideNumberingValue = new StartOverrideNumberingValue() { Val = startFrom }
            };
            numberingInstance.AppendChild(levelOverride);
        }


        return numberingInstance;
    }

    public static NumberingInstance AddBulletedListNumbering(this WordprocessingDocument document,
        AbstractNum? abstractNum = null)
    {
        var numbering = document.GetOrCreateNumbering();

        if (abstractNum == null)
        {
            var abstractNumberId = numbering.Elements<AbstractNum>().Count() + 1;

            abstractNum = new AbstractNum(
                new Level(
                    new NumberingFormat() { Val = NumberFormatValues.Bullet },
                    new LevelText() { Val = "·" }
                )
                { LevelIndex = 0 }
            )
            { AbstractNumberId = abstractNumberId };

            numbering.AddAbstractNumbering(abstractNum);
        }


        var numId = numbering.Elements<NumberingInstance>().Count() + 1;

        var numberingInstance = new NumberingInstance(
            new AbstractNumId() { Val = abstractNum.AbstractNumberId }
        )
        { NumberID = numId };

        numbering.AddNumberingInstance(numberingInstance);

        return numberingInstance;
    }

    public static void AddAbstractNumbering(this Numbering numbering, AbstractNum abstractNum)
    {
        numbering.InsertAfterLastOfType(abstractNum);
        numbering.Save();
    }

    public static void AddNumberingInstance(this Numbering numbering, NumberingInstance numberingInstance)
    {
        numbering.InsertAfterLastOfType(numberingInstance);
        numbering.Save();
    }

    public static Numbering GetOrCreateNumbering(this WordprocessingDocument document)
    {
        Debug.Assert(document.MainDocumentPart != null, "document.MainDocumentPart != null");

        if (document.MainDocumentPart.NumberingDefinitionsPart == null)
        {
            var part = document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
        }

        var numbering = document.MainDocumentPart.NumberingDefinitionsPart!.Numbering;
        return numbering;
    }

}
