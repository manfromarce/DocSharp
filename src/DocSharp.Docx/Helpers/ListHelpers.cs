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
    /// <summary>
    /// Helper function to retrieve Numbering part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static Numbering? GetNumberingPart(this OpenXmlElement element)
    {
        return element.GetMainDocumentPart()?.NumberingDefinitionsPart?.Numbering;
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

    public static Level? GetListLevel(Numbering numberingPart, int levelIndex, int numId, int abstractNumId)
    {
        var num = numberingPart.FirstOrDefault<NumberingInstance>(x => x.NumberID != null && x.NumberID == numId);
        var abstractNum = numberingPart.FirstOrDefault<AbstractNum>(x => x.AbstractNumberId != null && x.AbstractNumberId == abstractNumId);
        var level = abstractNum?.FirstOrDefault<Level>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
        var levelOverride = num?.FirstOrDefault<LevelOverride>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);

        // Use LevelOverride if present
        return levelOverride?.Level ?? level;
    }

    public static Level GetOrCreateListLevel(this ParagraphProperties pPr, MainDocumentPart mainPart)
    {
        pPr.NumberingProperties ??= new NumberingProperties();
        return pPr.NumberingProperties.GetOrCreateListLevel(mainPart);
    }

    public static Level GetOrCreateListLevel(this NumberingProperties numPr, MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.GetOrCreateNumbering();
        if (numPr.NumberingId == null || numPr.NumberingId.Val == null)
        {
           numPr.NumberingId = new NumberingId() { Val = numberingPart.GetAvailableNumberingInstanceId() };
        }
        int numberingId = numPr.NumberingId.Val;
        var num = numberingPart.FirstOrDefault<NumberingInstance>(x => x.NumberID != null && x.NumberID == numberingId) ?? 
                  numberingPart.AddNumberingInstance();

        if (num.AbstractNumId == null || num.AbstractNumId.Val == null)
        {
           num.AbstractNumId = new AbstractNumId() { Val = numberingPart.GetAvailableAbstractNumberId() };
        }
        var abstractNumId = num.AbstractNumId.Val;
        var abstractNum = numberingPart.FirstOrDefault<AbstractNum>(x => x.AbstractNumberId == abstractNumId) ??
                          numberingPart.AddAbstractNumbering(abstractNumId);

        int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;
        var level = abstractNum.FirstOrDefault<Level>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
        var levelOverride = num?.FirstOrDefault<LevelOverride>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
        // Use LevelOverride if present
        return levelOverride?.Level ?? level ?? abstractNum.AddLevel(levelIndex);
    }

    public static Level CreateListLevel(this ParagraphProperties pPr, MainDocumentPart mainPart)
    {
        pPr.NumberingProperties = new NumberingProperties();
        return pPr.NumberingProperties.CreateListLevel(mainPart);
    }

    public static Level CreateListLevel(this NumberingProperties numPr, MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.GetOrCreateNumbering();
        var abstractNumbering = numberingPart.AddAbstractNumbering();
        var numberingInstance = numberingPart.AddNumberingInstance();
        
        numberingInstance.AbstractNumId = new AbstractNumId() { Val = abstractNumbering.AbstractNumberId!.Value };
        var level = abstractNumbering.AddLevel(0);

        numPr.NumberingId = new NumberingId() { Val = numberingInstance.NumberID!.Value };
        numPr.NumberingLevelReference = new NumberingLevelReference() { Val = 0 };

        return level;
    }

    public static Level? GetListLevel(this ParagraphProperties? pPr)
    {
        return (pPr?.NumberingProperties).GetListLevel();
    }

    public static Level? GetListLevel(this NumberingProperties? numPr)
    {
        if (numPr == null)
        {
            return null;
        }

        var numberingPart = numPr.GetNumberingPart();
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int numberingId = numPr.NumberingId.Val;
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;

            var num = numberingPart.FirstOrDefault<NumberingInstance>(x => x.NumberID != null && x.NumberID == numberingId);
            var abstractNumId = num?.AbstractNumId?.Val;
            if (abstractNumId != null)
            {
                var abstractNum = numberingPart.FirstOrDefault<AbstractNum>(x => x.AbstractNumberId == abstractNumId);
                var level = abstractNum?.FirstOrDefault<Level>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
                var levelOverride = num?.FirstOrDefault<LevelOverride>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);

                // Use LevelOverride if present
                return levelOverride?.Level ?? level;
            }
        }
        return null;
    }

    public static int GetAvailableAbstractNumberId(this Numbering numbering)
    {
        int existingId = numbering.Elements<AbstractNum>().Max(x => x.AbstractNumberId ?? -1) ?? -1;
        return existingId + 1;
    }

    public static int GetAvailableNumberingInstanceId(this Numbering numbering)
    {
        int existingId = numbering.Elements<NumberingInstance>().Max(x => x.NumberID ?? 0) ?? 0;
        // Note that numbering instance id must start from 1; 0 does not work
        return existingId + 1;
    }

    public static int GetAvailableLevelId(this AbstractNum abstractNum)
    {
        int existingId = abstractNum.Elements<Level>().Max(x => x.LevelIndex ?? -1) ?? -1;
        return existingId + 1;
    }

    public static int GetAvailableLevelOverrideId(this NumberingInstance numberingInstance)
    {
        int existingId = numberingInstance.Elements<LevelOverride>().Max(x => x.LevelIndex ?? -1) ?? -1;
        return existingId + 1;
    }

    public static AbstractNum AddAbstractNumbering(this Numbering numbering)
    {
        return numbering.AddAbstractNumbering(numbering.GetAvailableAbstractNumberId());
    }

    public static AbstractNum AddAbstractNumbering(this Numbering numbering, int id)
    {
        return numbering.AddAbstractNumbering(new AbstractNum() { AbstractNumberId = id });
    }

    public static AbstractNum AddAbstractNumbering(this Numbering numbering, AbstractNum abstractNum)
    {
        numbering.InsertAfterLastOfType<AbstractNum>(abstractNum);
        return abstractNum;
    }

    public static NumberingInstance AddNumberingInstance(this Numbering numbering, AbstractNum linkedAbstractNumb)
    {
        linkedAbstractNumb.AbstractNumberId ??= numbering.GetAvailableAbstractNumberId();
        return numbering.AddNumberingInstance(linkedAbstractNumb.AbstractNumberId);
    }

    public static NumberingInstance AddNumberingInstance(this Numbering numbering, int abstractNumId)
    {        
        var numberingInstance = new NumberingInstance(new AbstractNumId() { Val = abstractNumId }) { NumberID = numbering.GetAvailableNumberingInstanceId() };
        return numbering.AddNumberingInstance(numberingInstance);
    }

    public static NumberingInstance AddNumberingInstance(this Numbering numbering, NumberingInstance numberingInstance)
    {
        numbering.InsertAfterLastOfType<NumberingInstance>(numberingInstance);
        return numberingInstance;
    }

    public static NumberingInstance AddNumberingInstance(this Numbering numbering)
    {
        return numbering.AddNumberingInstance(numbering.GetAvailableNumberingInstanceId());
    }

    public static NumberingInstance AddNumberingInstance(this AbstractNum abstractNum)
    {
        if (abstractNum.GetFirstAncestor<Numbering>() is Numbering numbering)
            return numbering.AddNumberingInstance(abstractNum);
        else return new NumberingInstance() { AbstractNumId = new AbstractNumId() { Val = abstractNum.AbstractNumberId }};
    }

    public static Level GetSimpleLevel(this AbstractNum abstractNum)
    {
        return abstractNum.Elements<Level>().LastOrDefault() ?? abstractNum.AddLevel();
    }

    public static Level AddLevel(this AbstractNum abstractNum)
    {
        return abstractNum.AddLevel(abstractNum.GetAvailableLevelId());
    }

    public static Level AddLevel(this AbstractNum abstractNum, int levelIndex)
    {
        abstractNum.RemoveAll<Level>(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
        return abstractNum.AppendChild(new Level() { LevelIndex = levelIndex });
    }

    public static Level AddLevel(this AbstractNum abstractNum, Level level)
    {
        return abstractNum.AppendChild(level);
    }

    public static LevelOverride AddLevelOverride(this NumberingInstance numberingInstance, LevelOverride lvlOverride)
    {
        return numberingInstance.AppendChild(lvlOverride);
    }

    internal static LevelOverride CreateLevelOverride(this NumberingInstance numberingInstance)
    {
        var id = numberingInstance.GetAvailableLevelOverrideId();
        return new LevelOverride(new Level() { LevelIndex = id }) { LevelIndex = id };
    }

    public static LevelOverride AddLevelOverride(this NumberingInstance numberingInstance)
    {
        return numberingInstance.AppendChild(numberingInstance.CreateLevelOverride());
    }

    public static Numbering GetOrCreateNumbering(this WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart ?? document.AddMainDocumentPart();
        return mainPart.GetOrCreateNumbering();
    }

    public static Numbering GetOrCreateNumbering(this MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        return numberingPart.Numbering;
    }

    public static Numbering? GetOrCreateNumbering(this OpenXmlElement element)
    {
        return element?.GetMainDocumentPart()?.GetOrCreateNumbering();
    }

    internal static void PrependLevelText(this Level level, string val)
    {
        level.LevelText ??= new LevelText();
        if (level.LevelText.Val != null)
            level.LevelText.Val = val + level.LevelText.Val;
        else level.LevelText.Val = val;
    }

    internal static void AppendLevelText(this Level level, string val)
    {
        level.LevelText ??= new LevelText();
        if (level.LevelText.Val != null)
            level.LevelText.Val = level.LevelText.Val + val;
        else level.LevelText.Val = val;
    }

    internal static void SetLevelText(this Level level, string val)
    {
        level.LevelText ??= new LevelText();
        level.LevelText.Val = val;
    }

    internal static string GetLevelText(this Level level)
    {
        return level.LevelText?.Val?.Value ?? string.Empty;
    }
}
