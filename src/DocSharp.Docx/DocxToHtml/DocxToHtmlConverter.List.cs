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
    private readonly Dictionary<(int NumberingId, int LevelIndex), int> _listLevelCounters = new();

    internal void ProcessListItem(NumberingProperties numPr, HtmlTextWriter sb)
    {
        // Note: we don't produce real HTML lists (<ul> / <ol>) because they are very limited compared to DOCX,
        // and preserving the original list format would be complicated. 
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
                var effectiveLevel = levelOverride?.Level ?? level;

                var start = levelOverride?.StartOverrideNumberingValue?.Val ?? effectiveLevel?.StartNumberingValue?.Val;
                var levelText = effectiveLevel?.LevelText?.Val;
                var listType = effectiveLevel?.NumberingFormat?.Val;
                var runPr = effectiveLevel?.NumberingSymbolRunProperties;

                if (effectiveLevel != null &&
                    listType != null &&
                    listType != NumberFormatValues.None)
                {
                    string listText = string.Empty;
                    var key = (NumberingId: numberingId, LevelIndex: levelIndex);

                    // Restart numbering
                    // var restart = effectiveLevel.LevelRestart?.Val;
                    // if (restart!= null && restart.HasValue && restart.Value <= levelIndex + 1)
                    // {
                    //     for (int i = restart.Value - 1; i <= levelIndex; i++)
                    //     {
                    //         var restartKey = (NumberingId: numberingId, LevelIndex: i);
                    //         _listLevelCounters[restartKey] = start ?? 1;
                    //     }
                    // }
                    // else
                    // {
                    if (!_listLevelCounters.ContainsKey(key))
                    {
                        _listLevelCounters[key] = start ?? 1;
                    }
                    else
                    {
                        _listLevelCounters[key]++;
                    }
                    // }

                    // Reset counters for deeper levels of this NumberingId
                    foreach (var deeperLevel in _listLevelCounters.Keys
                        .Where(k => k.NumberingId == numberingId && k.LevelIndex > levelIndex)
                        .ToList())
                    {
                        _listLevelCounters.Remove(deeperLevel);
                    }

                    if (listType == NumberFormatValues.Bullet)
                    {
                        if (levelText?.Value != null)
                        {
                            listText += levelText.Value;
                        }
                        else
                        {
                            listText += "•";
                        }
                    }
                    else
                    {
                        // Numbered list
                        listText += ListHelpers.GetNumberString(levelText, listType, numberingId, levelIndex, _listLevelCounters);
                    }

                    var levelSuffix = effectiveLevel.LevelSuffix?.Val;
                    if (levelSuffix == null || levelSuffix.Value == LevelSuffixValues.Tab)
                    {
                        listText += "\u2001"; // quad space (&#x2001;)
                    }
                    else if (levelSuffix.Value == LevelSuffixValues.Space)
                    {
                        listText += '\u00A0'; // non-breaking space (translated to &emsp;)
                    }

                    if (runPr != null)
                    {
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
