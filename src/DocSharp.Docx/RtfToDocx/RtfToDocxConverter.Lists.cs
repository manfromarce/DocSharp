using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private void ParseListTable(RtfDestination dest)
    {
        
    }

    private void ParseListOverrideTable(RtfDestination dest)
    {
        
    }

    private Level EnsureLevel()
    {
        currentLevel ??= pPr.GetOrCreateListLevel(mainPart);
        return currentLevel;
    }

    private bool ProcessLegacyListControlWord(RtfControlWord cw)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "pnlvl": // Paragraph level, where the value is a level from 1 to 9.
                if (cw.HasValue && cw.Value!.Value > 0)
                {
                    EnsureLevel().LevelIndex = cw.Value!.Value - 1; // in Open XML it starts from 0 instead
                }
                return true;
            // case "pnlvlblt": // Bulleted paragraph (corresponds to level 11). The actual character used for the bullet is stored in the \pntxtb group.
            // Handled by RtfNumberFormatMapper.GetNumberFormat
            case "pnlvlbody": // Simple paragraph numbering (corresponds to level 10).
                EnsureLevel().NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                // (will be overwritten later by more specific number format, if specified).
                return true;
            case "pnlvlcont": // Continue numbering but do not display number ("skip numbering")
                // (unclear if the level text or even the whole paragraph should be hidden, 
                // for now set NumberingFormat to None to hide the number)
                EnsureLevel().NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.None };
                return true;

            // These cannot be set on individual list items in DOCX; ignore for now.
            case "pnnumonce": // Number each cell only once in a table (default is to number each paragraph in a table).
                return true;
            case "pnacross": // Number across rows (default is to number down columns).
                return true;
            case "pnhang": // Paragraph uses a hanging indent
                return true;
            case "pnrestart": // Restart numbering after each section break. This control word is used only in conjunction with the Heading Numbering feature (applying multilevel numbering to Heading style definitions)
                return true;

            case "pnstart": // Start at number
                if (cw.HasValue)
                {
                    EnsureLevel().StartNumberingValue = new StartNumberingValue() { Val = cw.Value!.Value };
                }
                return true;

            case "pnindent": // Minimum distance from margin to body text.
                if (cw.HasValue)
                {
                    EnsureLevel().LegacyNumbering = new LegacyNumbering
                    {
                        Legacy = true,
                        LegacyIndent = cw.Value!.Value.ToStringInvariant()
                    };
                }
                return true;
            case "pnsp": // Distance from number text to body text.
                if (cw.HasValue)
                {
                    EnsureLevel().LegacyNumbering = new LegacyNumbering
                    {
                        Legacy = true,
                        LegacySpace = cw.Value!.Value.ToStringInvariant()
                    };
                }
                return true;

            case "pnprev": // Used for multilevel lists. Include information from previous level in this level; for example, 1, 1.1, 1.1.1, 1.1.1.1.
                // Unclear how it works in RTF (should we replace %1 with %1.%2 and so on in the level text ?); ignore for now
                return true;
                
            // Set the following properties on AbstractNumId --> Level --> NumberingSymbolRunProperties
            case "pnf": // Font number for number/bullet
                if (cw.HasValue)
                {
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    if (fontTable.TryGetValue(cw.Value!.Value, out var fname) && !string.IsNullOrEmpty(fname))
                    level.NumberingSymbolRunProperties.RunFonts = new RunFonts() { Ascii = fname, HighAnsi = fname, EastAsia = fname, ComplexScript = fname };
                }
                return true;
            case "pnfs": // Font size
                if (cw.HasValue)
                {
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.FontSize = new FontSize() { Val = cw.Value!.Value.ToStringInvariant() };
                }
                return true;
            case "pncf": // Font color
                if (cw.HasValue)
                {
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    var idx = cw.Value!.Value;
                    if (idx >= 0 && idx < colorTable.Count)
                    {
                        var c = colorTable[idx];
                        var hex = c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
                        level.NumberingSymbolRunProperties.Color = new Color() { Val = hex };
                    }
                }
                return true;
            case "pnb": // Bullet/number is bold
                if (cw.HasValue && cw.Value == 0)
                {
                    // Disabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Bold = new Bold() { Val = false };
                }
                else
                {
                    // Enabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Bold = new Bold() { Val = true };
                }
                return true;
            case "pni": // Bullet/number is italic
                if (cw.HasValue && cw.Value == 0)
                {
                    // Disabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Italic = new Italic() { Val = false };
                }
                else
                {
                    // Enabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Italic = new Italic() { Val = true };
                }
                return true;
            case "pnul": // Bullet/number is underlined
                if (cw.HasValue && cw.Value == 0)
                {
                    // Disabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Underline = new Underline() { Val = UnderlineValues.None };
                }
                else
                {
                    // Enabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Underline = new Underline() { Val = UnderlineValues.Single };
                }
                return true;
            case "pnulw": // Bullet/number has word underline
                if (cw.HasValue && cw.Value == 0)
                {
                    // Disabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Underline = new Underline() { Val = UnderlineValues.None };
                }
                else
                {
                    // Enabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Underline = new Underline() { Val = UnderlineValues.Words };
                }
                return true;
             case "pnstrike": // Bullet/number is strikethrough
                if (cw.HasValue && cw.Value == 0)
                {
                    // Disabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Strike = new Strike() { Val = false };
                }
                else
                {
                    // Enabled
                    var level = EnsureLevel();
                    level.NumberingSymbolRunProperties ??= new NumberingSymbolRunProperties();
                    level.NumberingSymbolRunProperties.Strike = new Strike() { Val = true };
                }
                return true;

            // Set alignment on AbstractNumId --> Level --> LevelJustification
            case "pnqc": // Centered numbering
                EnsureLevel().LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Center };
                return true;
            case "pnql": // Left-aligned numberings
                EnsureLevel().LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };
                return true;
            case "pnqr": // Right-justified numbering
                EnsureLevel().LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Right };
                return true;

            default: 
                // Map \pndec, \pnucltr, \pnlcltr, ... to Open XML NumberFormatValues
                if (name.StartsWith("pn") && RtfNumberFormatMapper.GetNumberFormat(name) is NumberFormatValues numberFormat)
                {
                    var level = EnsureLevel();
                    // Set numbering format for the level
                    level.NumberingFormat = new NumberingFormat() { Val = numberFormat };

                    // If the format is NumberFormatValues.Bullet or NumberFormatValues.None, remove "%1" from the level text.
                    // (it may have been already inserted if \pntxta or \pntxtb is found before the numbering format is determined). 
                    var text = level.GetLevelText();
                    if (numberFormat == NumberFormatValues.Bullet || numberFormat == NumberFormatValues.None)
                    {
                        if (text.Contains("%1"))
                            level.SetLevelText(text.Replace("%1", ""));
                    }
                    else if (text == string.Empty)
                        // Otherwise if the the number format is not bullet/none (the list is numbered), 
                        // and the level text has not been set (\pntxta or \pntxtb have not been found yet), 
                        // set the level text to %1 (it will be replaced by the actual number by word processors).
                        level.SetLevelText("%1");
                    
                    return true;
                }
                break;
        }
        return false;
    }

    private Numbering EnsureNumberingPart()
    {
        numberingDefinitionsPart ??= mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingDefinitionsPart.Numbering ??= new Numbering();
        return numberingDefinitionsPart.Numbering;
    }
}