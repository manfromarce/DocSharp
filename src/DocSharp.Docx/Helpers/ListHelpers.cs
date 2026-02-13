using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class ListHelpers
{
    public static Level? GetListLevel(Numbering numberingPart, int levelIndex, int numId, int abstractNumId)
    {
        var num = numberingPart.Elements<NumberingInstance>()
                                .FirstOrDefault(x => x.NumberID != null &&
                                                     x.NumberID == numId);
        var abstractNum = numberingPart.Elements<AbstractNum>()
                        .FirstOrDefault(x => x.AbstractNumberId != null && 
                                             x.AbstractNumberId == abstractNumId);
        var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                       x.LevelIndex == levelIndex);
        var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null &&
                                                                               x.LevelIndex == levelIndex);

        // Use LevelOverride if present
        return levelOverride?.Level ?? level;
    }

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

    public static string GetNumberString(string? levelText, NumberingFormat? numberingFormat, Dictionary<int, (int numId, int numberingId, int counter)> listLevelCounters, CultureInfo? culture = null)
    {
        var listType = numberingFormat?.Val ?? NumberFormatValues.Decimal; // if not specified it should be assumed decimal (regular numbered list)
        if (listType == NumberFormatValues.Bullet || string.IsNullOrEmpty(levelText))
        {
            return "•"; // // Bullet text and font is handled separately
        }

        string formattedText = levelText!;
        // TODO: use the document language; InvariantCulture or CurrentCulture by default
        culture ??= CultureInfo.InvariantCulture;

        // Retrieve the max level requested by level text
        int maxPlaceholder = 0;
        foreach (Match match in Regex.Matches(levelText, @"%(\d+)"))
        {
            if (match.Groups.Count > 1 &&
                int.TryParse(match.Groups[1].Value, out int n) && n > maxPlaceholder)
                maxPlaceholder = n;
        }

        // For each placeholder in the level text, search for the appropriate counter in the level hierarchy
        for (int i = 1; i <= maxPlaceholder; i++)
        {
            int value = 1;
            // Try to find the value for this level (note: levels use 0-based index, placeholders use 1-based index)
            if (listLevelCounters.TryGetValue(i - 1, out (int numId, int abstractNumId, int counter) tuple))
            {
                value = tuple.counter;
            }

            // Format the value depending on the list type
            string replacement;
            if (listType == NumberFormatValues.LowerLetter)
            {
                replacement = NumberToLetter(value, false);
            }
            else if (listType == NumberFormatValues.UpperLetter)
            {
                replacement = NumberToLetter(value, true);
            }
            else if (listType == NumberFormatValues.LowerRoman)
            {
                replacement = NumberToRomanLetter(value, false);
            }
            else if (listType == NumberFormatValues.UpperRoman)
            {
                replacement = NumberToRomanLetter(value, true);
            }
            else if (listType == NumberFormatValues.DecimalEnclosedCircle ||
                        listType == NumberFormatValues.DecimalEnclosedCircleChinese)
            {
                replacement = NumberToCircledNumber(value, true);
            }
            else if (listType == NumberFormatValues.Chicago)
            {
                replacement = NumberToChicago(value);
            }
            else if (listType == NumberFormatValues.Hex)
            {
                replacement = value.ToString("X", CultureInfo.InvariantCulture); // Always use invariant culture for hex format
            }
            else if (listType == NumberFormatValues.NumberInDash)
            {
                replacement = $"- {value.ToString(culture)} -";
            }
            else if (listType == NumberFormatValues.DecimalEnclosedFullstop)
            {
                replacement = $"{value.ToString(culture)}.";
            }
            else if (listType == NumberFormatValues.DecimalEnclosedParen)
            {
                replacement = $"({value.ToString(culture)})";
            }
            else if (listType == NumberFormatValues.DecimalZero)
            {
                replacement = value.ToString("00");
            }
            else if (listType == NumberFormatValues.Custom &&
                     numberingFormat?.Format?.Value != null)
            {
                switch (numberingFormat.Format.Value)
                {
                    // These are few standard formats created by Microsoft Word
                    case "01, 02, 03, ...":
                        replacement = value.ToString("00");
                        break;
                    case "001, 002, 003, ...":
                        replacement = value.ToString("000");
                        break;
                    case "0001, 0002, 0003, ...":
                        replacement = value.ToString("0000");
                        break;
                    case "00001, 00002, 00003, ...":
                        replacement = value.ToString("00000");
                        break;
                    default:
                        replacement = value.ToString(culture);
                        break;
                }
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
                replacement = "•"; // Bullet text and font is handled separately
            }
            else if (listType == NumberFormatValues.None)
            {
                replacement = string.Empty;
            }
            else
            {
                // Regular number
                replacement = value.ToString(culture);
            }

            // Replace placeholder with the formatted value
            formattedText = formattedText.Replace($"%{i}", replacement);
        }

        return formattedText;
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
        numbering.InsertAfterLastOfType<AbstractNum>(abstractNum);
        numbering.Save();
    }

    public static void AddNumberingInstance(this Numbering numbering, NumberingInstance numberingInstance)
    {
        numbering.InsertAfterLastOfType<NumberingInstance>(numberingInstance);
        numbering.Save();
    }

    public static Numbering GetOrCreateNumbering(this WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart ?? document.AddMainDocumentPart();
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        return numberingPart.Numbering;
    }

}
