using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextWriterBase<HtmlTextWriter>
{
    private readonly Dictionary<int, (int numId, int abstractNumId, int counter)> _listLevelCounters = new();

    internal void ProcessListItem(NumberingProperties numPr, HtmlTextWriter sb, bool isHidden = false)
    {
        // Note: we don't produce real HTML lists (<ul> / <ol>) because they are very limited compared to DOCX,
        // and preserving the original list format would be complicated. 
        var numberingPart = OpenXmlHelpers.GetNumberingPart(numPr);
        if (numberingPart != null && numPr.NumberingId?.Val != null)
        {
            int numberingId = numPr.NumberingId.Val;
            int levelIndex = numPr.NumberingLevelReference?.Val ?? 0;

            // Find the NumberingInstance, AbstractNumbering and level associated with this list item
            var num = numberingPart.Elements<NumberingInstance>()
                                   .FirstOrDefault(x => x.NumberID != null &&
                                                        x.NumberID == numberingId);
            if (num?.AbstractNumId?.Val != null)
            {
                var abstractNumId = num.AbstractNumId.Val.Value;
                var abstractNum = numberingPart.Elements<AbstractNum>()
                                  .FirstOrDefault(x => x.AbstractNumberId != null && x.AbstractNumberId == abstractNumId);
                
                var level = abstractNum?.Elements<Level>().FirstOrDefault(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
                var levelOverride = num?.Elements<LevelOverride>().FirstOrDefault(x => x.LevelIndex != null && x.LevelIndex == levelIndex);
                // Use LevelOverride if present
                var effectiveLevel = levelOverride?.Level ?? level;

                // Retrieve the level start number, text and format.
                var start = 0; // if not specified it should be assumed 0, not 1 (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.StartNumberingValue?view=openxml-3.0.1)
                if (levelOverride?.StartOverrideNumberingValue?.Val != null)
                    start = levelOverride.StartOverrideNumberingValue.Val.Value;
                else if (effectiveLevel?.StartNumberingValue?.Val != null)
                    start = effectiveLevel.StartNumberingValue.Val.Value;
                var levelText = effectiveLevel?.LevelText?.Val;
                var listType = effectiveLevel?.NumberingFormat?.Val ?? NumberFormatValues.Decimal; // if not specified it should be assumed decimal (regular numbered list)
                                                                                                   // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.NumberingFormat?view=openxml-3.0.1)
                var runPr = effectiveLevel?.NumberingSymbolRunProperties;

                if (effectiveLevel != null && listType != NumberFormatValues.None)
                {
                    // The dictionary will contain at maximum 9 levels
                    if (_listLevelCounters.ContainsKey(levelIndex))
                    {
                        // If the current level index is already in the dictionary, check its abstract numbering ID.
                        var state = _listLevelCounters[levelIndex];
                        if (state.abstractNumId != abstractNumId)
                        {
                            // If the AbstractNumId is different, restart the level from its start value.
                            _listLevelCounters.Remove(levelIndex);
                            _listLevelCounters.Add(levelIndex, (numberingId, abstractNumId, start));
                        }
                        else
                        {
                            // If the AbstractNumId is the same, continue numbering.
                            int last = state.counter;
                            _listLevelCounters.Remove(levelIndex);
                            _listLevelCounters.Add(levelIndex, (numberingId, abstractNumId, last + 1));
                        }
                    }
                    else
                    {
                        // If the dictionary does not contain this level, start the level from its start value.
                        _listLevelCounters.Add(levelIndex, (numberingId, abstractNumId, start));
                    }

                    // Reset counters for deeper levels, to avoid continue numbering
                    foreach (var lvlIndex in _listLevelCounters.Keys
                                                .Where(x => x > levelIndex) // filter levels with an higher index than the current
                                                .ToList())
                    {
                        // By default, a level restarts from the start value each time the previous level is used, e.g.:
                        // 1
                        //    a
                        //    b
                        // 2 
                        //    a (does not continue the previous nested list numbering)
                        // However, this can be overriden by the LevelRestart value, which must still be minor than the current level.
                        // A level restart value of 0 means the level should never restart.
                        // (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.levelrestart?view=openxml-3.0.1) 

                        var state = _listLevelCounters[lvlIndex];
                        Level? deeperLevel = ListHelpers.GetListLevel(numberingPart, lvlIndex, state.numId, state.abstractNumId);
                        // Try to get the levelRestart element (note: its value has 1-based index like placeholders, while level index starts from 0)
                        var levelRestart = deeperLevel?.LevelRestart?.Val != null ? Math.Min(lvlIndex, deeperLevel.LevelRestart.Val.Value - 1) : lvlIndex;
                        // (if levelRestart is not present, uses the current level index by default (see explanation above).
                        levelRestart = Math.Max(levelRestart, 0);
                        if (levelIndex < levelRestart)
                        {
                            // Remove counter for deeper levels depending on level indexes and levelRestart (if specified). 
                            _listLevelCounters.Remove(lvlIndex);
                        }
                    }

                    if (!isHidden)
                    {
                        string listText;
                        if (listType == NumberFormatValues.Bullet)
                        {
                            // For bulleted lists, level text can be returned as-is.
                            listText = levelText?.Value != null ? levelText.Value : "•";
                        }
                        else
                        {
                            // For numbered lists, get the number text depending on the list format and level counters.
                            listText = ListHelpers.GetNumberString(levelText, listType, _listLevelCounters);
                        }

                        // Add the suffix
                        var levelSuffix = effectiveLevel.LevelSuffix?.Val;
                        if (levelSuffix == null || levelSuffix.Value == LevelSuffixValues.Tab)
                        {
                            listText += "\u2001"; // quad space (&#x2001;)
                        }
                        else if (levelSuffix.Value == LevelSuffixValues.Space)
                        {
                            listText += '\u00A0'; // non-breaking space (&emsp;)
                        }

                        if (runPr != null)
                        {
                            // Process formatting for the number/bullet
                            var rPr = new RunProperties();
                            foreach (var runProperty in runPr)
                            {
                                rPr.AppendChild(runProperty.CloneNode(true));
                            }
                            ProcessRun(new Run(rPr, new Text(listText)), sb);
                        }
                        else
                        {
                            ProcessRun(new Run(new Text(listText)), sb);
                        }
                    }
                }
            }
        }
    }
}
