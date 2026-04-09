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
    private Dictionary<int, int> rtfListTableMap = new();
    private Dictionary<int, int> rtfListOverrideMap = new();

    private void ParseListTable(RtfDestination dest)
    {
        // Ensure numbering part exists
        var numbering = mainPart.GetOrCreateNumbering();
        foreach (var token in dest.Tokens)
        {
            if (token is RtfDestination listDest && string.Equals(listDest.Name, "list", StringComparison.OrdinalIgnoreCase))
            {
                // Create an AbstractNum
                var abstractNum = numbering.AddAbstractNumbering();
                                
                int? listId = null;
                foreach (var lt in listDest.Tokens)
                {
                    if (lt is RtfControlWord cw)
                    {
                        string name = cw.Name.ToLowerInvariant();
                        switch (name)
                        {
                            case "listid":
                                if (cw.HasValue)
                                    listId = cw.Value!.Value; 
                                break;
                            case "listsimple":
                                if (cw.HasValue && cw.Value!.Value == 1)
                                    abstractNum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.SingleLevel };
                                else 
                                    abstractNum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
                                break;
                            case "listhybrid":
                                abstractNum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
                                break;
                            // case "listrestarthdn"
                            // Restart at each section (for Word 95 compatibility only) (not available in DOCX)
                        }
                    }
                    else if (lt is RtfDestination subDest)                    
                    {
                        if (string.Equals(subDest.Name, "listlevel", StringComparison.OrdinalIgnoreCase))
                        {
                            // Add level with default values
                            var lvl = abstractNum.AddLevel();
                            lvl.StartNumberingValue = new StartNumberingValue() { Val = 1 };
                            lvl.NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                            lvl.LevelText = new LevelText() { Val = "%1." };

                            // Map level control words
                            ProcessLevel(subDest, ref lvl);
                        }                        
                    }
                }

                // Check if the produced list is valid
                if (abstractNum.Elements<Level>().Any() && listId.HasValue)
                {
                    // Map RTF list id to AbstractNumId
                    if (abstractNum.AbstractNumberId != null && rtfListTableMap != null)
                    {
                        rtfListTableMap[listId.Value] = abstractNum.AbstractNumberId.Value;
                    }
                }
            }
            else if (token is RtfDestination listPictureDest && string.Equals(listPictureDest.Name, "listpicture", StringComparison.OrdinalIgnoreCase))
            {
                // Search for inner \shppict or pict groups
                int pictId = 0;
                foreach (var pictGroup in listPictureDest.Tokens.OfType<RtfDestination>()
                                                                .Where(d => string.Equals(d.Name, "shppict", StringComparison.OrdinalIgnoreCase) 
                                                                         || string.Equals(d.Name, "pict", StringComparison.OrdinalIgnoreCase)))
                {
                    // Get the pict group
                    var pict = pictGroup.Name.Equals("pict", StringComparison.OrdinalIgnoreCase) 
                                ? pictGroup 
                                : pictGroup.Tokens.OfType<RtfDestination>().FirstOrDefault(g => string.Equals(g.Name, "pict", StringComparison.OrdinalIgnoreCase));
                    if (pict != null)
                    {
                        var pictureBullet = ProcessPicture<PictureBulletBase>(pict, isPictureBullet: true);
                        if (pictureBullet != null)
                            numbering.Append(new NumberingPictureBullet(pictureBullet)
                            {
                                NumberingPictureBulletId = pictId
                            });
                    }
                    ++pictId;
                }
            }
        }
    }

    private void ParseListOverrideTable(RtfDestination dest)
    {
        // Ensure numbering part exists
        var numbering = mainPart.GetOrCreateNumbering();
        foreach (var token in dest.Tokens)
        {
            if (token is RtfDestination ovGroup && string.Equals(ovGroup.Name, "listoverride", StringComparison.OrdinalIgnoreCase))
            {
                int? listId = null;
                int? ls = null;

                foreach (var t in ovGroup.Tokens)
                {
                    if (t is RtfControlWord cw)
                    {
                        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
                        if (name == "listid" && cw.HasValue)
                            listId = cw.Value;
                        else if (name == "ls" && cw.HasValue)
                            ls = cw.Value;
                    }
                }

                if (!listId.HasValue || !ls.HasValue)
                    continue;

                // Resolve abstract num id from earlier parsed listtable
                if (!rtfListTableMap.TryGetValue(listId.Value, out var abstractNumId))
                    continue;

                // Create a NumberingInstance linked to the abstractNumId
                var numberingInstance = numbering.AddNumberingInstance(abstractNumId);

                // Map RTF \ls value to the created numbering instance id
                if (numberingInstance.NumberID != null)
                {
                    rtfListOverrideMap[ls.Value] = numberingInstance.NumberID.Value;
                }

                // Parse level overrides (\lfolevel groups) inside this listoverride
                foreach (var inner in ovGroup.Tokens)
                {
                    if (inner is RtfDestination lfo && string.Equals(lfo.Name, "lfolevel", StringComparison.OrdinalIgnoreCase))
                    {
                        // Create a LevelOverride attached to the numbering instance
                        var levelOverride = numberingInstance.AddLevelOverride();
                        var targetLevel = levelOverride.Level; // the Level child inside the override

                        if (targetLevel == null)
                            continue;

                        // Inspect tokens inside \lfolevel
                        foreach (var it in lfo.Tokens)
                        {
                            if (it is RtfControlWord icw && icw.Name != null)
                            {
                                var iname = icw.Name.ToLowerInvariant();
                                if (string.Equals(iname, "listoverridestartat", StringComparison.OrdinalIgnoreCase) && icw.HasValue)
                                {
                                    levelOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue() { Val = icw.Value!.Value };
                                }
                                // Other simple overrides could be handled here if present
                            }
                            else if (it is RtfDestination g && string.Equals(g.Name, "listlevel", StringComparison.OrdinalIgnoreCase))
                            {
                                // Parse the level definition
                                ProcessLevel(g, ref targetLevel);
                            }
                        }
                    }
                }
            }
        }
    }

    private void ProcessLevel(RtfDestination dest, ref Level level)
    {
        var previousBorder = currentBorder;
        currentBorder = null;
        var levelParaPr = new ParagraphProperties();

        // Use a clean run state for the level
        fmtStack.Push(new FormattingState());

        foreach (var lt in dest.Tokens)
        {
            if (lt is RtfControlWord lcw && lcw.Name != null)
            {
                var name = lcw.Name.ToLowerInvariant();
                switch(name)
                {
                    case "levelnfc":
                    case "levelnfcn":
                        if (lcw.HasValue)
                        {
                            var format = RtfNumberFormatMapper.GetNumberFormat(name + lcw.Value!.Value.ToStringInvariant());
                            if (format != null)
                                level.NumberingFormat = new NumberingFormat() { Val = format };
                        }
                        break;
                    case "levelstartat":
                        level.StartNumberingValue = new StartNumberingValue() { Val = lcw.Value!.Value }; break;
                    case "levelfollow":
                        if (lcw.HasValue && lcw.Value == 0)
                            level.LevelSuffix = new LevelSuffix() { Val = LevelSuffixValues.Tab };
                        else if (lcw.HasValue && lcw.Value == 1)
                            level.LevelSuffix = new LevelSuffix() { Val = LevelSuffixValues.Space };
                        else 
                            level.LevelSuffix = new LevelSuffix() { Val = LevelSuffixValues.Nothing };                            
                        break;
                    case "leveljc":
                    case "leveljcn":
                        if (lcw.HasValue)
                        {
                            if (lcw.Value == 0) level.LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };
                            else if (lcw.Value == 1) level.LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Center };
                            else if (lcw.Value == 2) level.LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Right };
                        }
                        break;
                    case "levelpicture":
                        level.LevelPictureBulletId = new LevelPictureBulletId() { Val = lcw.Value!.Value }; break;
                    case "lvltentative":
                        if (!lcw.HasValue || lcw.Value!.Value == 1) 
                            level.Tentative = true; 
                        break;
                    case "levelnorestart": // Do not restart numbering at this level
                        if (!lcw.HasValue || lcw.Value!.Value == 1) 
                            level.LevelRestart = new LevelRestart() { Val = 0 };
                        break;
                    case "levellegal": 
                        if (!lcw.HasValue || lcw.Value!.Value == 1) 
                            level.IsLegalNumberingStyle = new IsLegalNumberingStyle() { Val = true};
                        break;

                    // Keywords emitted for compatibility with Word 95 or Word 6.0 only
                    case "levelold":
                        if (!lcw.HasValue || lcw.Value!.Value == 1)
                        {
                            level.LegacyNumbering ??= new LegacyNumbering();
                            level.LegacyNumbering.Legacy = true;
                        }
                        break;
                    case "levelindent":
                        if (lcw.HasValue)
                        {
                            level.LegacyNumbering ??= new LegacyNumbering();
                            level.LegacyNumbering.LegacyIndent = lcw.Value!.Value.ToStringInvariant();
                        }
                        break;
                    case "levelspace":
                        if (lcw.HasValue)
                        {
                            level.LegacyNumbering ??= new LegacyNumbering();
                            level.LegacyNumbering.LegacySpace = lcw.Value!.Value.ToStringInvariant();
                        }
                        break;
                    // case "levelprev":
                    // case "levelprevspace":
                    // Not available in DOCX

                    // case "levelpicturenosize":
                    
                    default: 
                        if (ProcessRunControlWord(lcw, TryPeek(fmtStack)))
                        {
                            break;
                        }
                        if (ProcessParagraphControlWord(lcw, levelParaPr))
                        {
                            break;
                        }
                        break;
                }
            }            
            else if (lt is RtfDestination dest2 && dest2.Name != null)
            {
                var dname = dest2.Name.ToLowerInvariant();
                if (dname == "leveltext")
                {
                    var sb = new StringBuilder();
                    ConvertGroupAsText(dest2, sb, true);
                    level.LevelText = new LevelText() { Val = sb.ToString().TrimEnd(";") };
                }                               
            }
        }
    
        currentBorder = previousBorder;

        // Note: we can't cast ParagraphProperties to PreviousParagraphProperties, 
        // but adding them directly should work as they have the same XML name and structure.
        if (levelParaPr.HasChildren)
            level.Append(levelParaPr);

        var generatedRun = CreateRunWithProperties(TryPeek(fmtStack));
        var rp = generatedRun.GetFirstChild<RunProperties>();
        if (rp != null && rp.HasChildren)
            // Note: we can't cast RunProperties to NumberingSymbolRunProperties, 
            // but adding them directly should work as they have the same XML name and structure.
            level.Append(rp.CloneNode(true));
    }
}