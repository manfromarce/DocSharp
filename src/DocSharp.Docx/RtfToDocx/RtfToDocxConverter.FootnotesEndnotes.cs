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
    internal FootnotesEndnotesType FootnotesEndnotes = FootnotesEndnotesType.FootnotesOnlyOrNothing;

    private bool ProcessFootnoteEndnoteControlWord(RtfControlWord cw, FormattingState runState)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch(name)
        {
            // Run level
            case "chftn": 
                if (container is Footnote || container.GetFirstAncestor<Footnote>() != null)
                    CreateRun().Append(new FootnoteReferenceMark());                    
                else if (container is Endnote || container.GetFirstAncestor<Endnote>() != null)
                    CreateRun().Append(new EndnoteReferenceMark());                    
                else
                    pendingFootnoteEndnoteRef = true;
                return true;
            case "chftnsep": 
                CreateRun().Append(new SeparatorMark());
                return true;
            case "chftnsepc":
                CreateRun().Append(new ContinuationSeparatorMark());
                return true;

            // Section level
            case "endnhere":
                // We don't need any special handling in DOCX.
                return true;
            case "sftnbj":
                EnsureSectionProperty<FootnoteProperties>().FootnotePosition = new FootnotePosition() { Val = FootnotePositionValues.PageBottom };
                return true;
            case "sftntj":
                EnsureSectionProperty<FootnoteProperties>().FootnotePosition = new FootnotePosition() { Val = FootnotePositionValues.BeneathText };
                return true;
            case "sftnrestart":
                EnsureSectionProperty<FootnoteProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.EachSection };
                return true;
            case "sftnrstcont":
                EnsureSectionProperty<FootnoteProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.Continuous };
                return true;
            case "sftnrstpg":
                EnsureSectionProperty<FootnoteProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.EachPage };
                return true;
            case "sftnstart":
                if (cw.HasValue)
                    EnsureSectionProperty<FootnoteProperties>().NumberingStart = new NumberingStart() { Val = (ushort)cw.Value!.Value };
                return true;
            case "saftnrestart":
                EnsureSectionProperty<EndnoteProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.EachSection };
                return true;
            case "saftnrstcont":
                EnsureSectionProperty<EndnoteProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.Continuous };
                return true;
            case "saftnstart":
                if (cw.HasValue)
                    EnsureSectionProperty<EndnoteProperties>().NumberingStart = new NumberingStart() { Val = (ushort)cw.Value!.Value };
                return true;

            // Document level
            case "ftnbj":
                EnsureSetting<FootnoteDocumentWideProperties>().FootnotePosition = new FootnotePosition() { Val = FootnotePositionValues.PageBottom };
                return true;
            case "ftntj":
                EnsureSetting<FootnoteDocumentWideProperties>().FootnotePosition = new FootnotePosition() { Val = FootnotePositionValues.BeneathText };
                return true;
            case "enddoc":
                if (FootnotesEndnotes == FootnotesEndnotesType.EndnotesOnly)
                    // In this case enddoc is emitted for backward compatibility only and is preceded/followed by aenddoc
                    return true;                    
                else
                    // Footnotes at document end are not available in Open XML, fallback to footnotes at section end
                    EnsureSetting<FootnoteDocumentWideProperties>().FootnotePosition = new FootnotePosition() { Val = FootnotePositionValues.SectionEnd };
                return true;
            case "endnotes":
                EnsureSetting<FootnoteDocumentWideProperties>().FootnotePosition = new FootnotePosition() { Val = FootnotePositionValues.SectionEnd };
                return true;
            case "ftnrestart":
                EnsureSetting<FootnoteDocumentWideProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.EachSection };
                return true;
            case "ftnrstpg":
                EnsureSetting<FootnoteDocumentWideProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.EachPage };
                return true;
            case "ftnrstcont":
                EnsureSetting<FootnoteDocumentWideProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.Continuous };
                return true;
            case "ftnstart":
                if (cw.HasValue)
                    EnsureSetting<FootnoteDocumentWideProperties>().NumberingStart = new NumberingStart() { Val = (ushort)cw.Value!.Value };
                return true;
            case "aenddoc":
                EnsureSetting<EndnoteDocumentWideProperties>().EndnotePosition = new EndnotePosition() { Val = EndnotePositionValues.DocumentEnd };
                return true;
            case "aendnotes":
                EnsureSetting<EndnoteDocumentWideProperties>().EndnotePosition = new EndnotePosition() { Val = EndnotePositionValues.SectionEnd };
                return true;
            case "aftnbj":
                // Endnotes at the bottom of page and beneath text are not available in Open XML. 
                // If only endnotes are enabled we could handle them as footnotes (in case aftnbj or aftntj is found), 
                // but that would require converting other control words too.
                // For now, just fallback to section end. 
                EnsureSetting<EndnoteDocumentWideProperties>().EndnotePosition = new EndnotePosition() { Val = EndnotePositionValues.SectionEnd };
                return true;
            case "aftntj":
                EnsureSetting<EndnoteDocumentWideProperties>().EndnotePosition = new EndnotePosition() { Val = EndnotePositionValues.SectionEnd };
                return true;
            case "aftnrestart":
                EnsureSetting<EndnoteDocumentWideProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.EachSection };
                return true;
            case "aftnrstcont":
                EnsureSetting<EndnoteDocumentWideProperties>().NumberingRestart = new NumberingRestart() { Val = RestartNumberValues.Continuous };
                return true;
            case "aftnstart":
                if (cw.HasValue)
                    EnsureSetting<EndnoteDocumentWideProperties>().NumberingStart = new NumberingStart() { Val = (ushort)cw.Value!.Value };
                return true;
            
            case "fet":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                        FootnotesEndnotes = FootnotesEndnotesType.FootnotesOnlyOrNothing;
                    else if (cw.Value == 1)
                        FootnotesEndnotes = FootnotesEndnotesType.EndnotesOnly;
                    else if (cw.Value == 2)
                        FootnotesEndnotes = FootnotesEndnotesType.Both;
                }
                return true;
        }

        if (cw.Name?.StartsWith("ftnn") == true)
        {
            var format = RtfNumberFormatMapper.GetNumberFormat(cw.Name);
            if (format != null)
            {
                settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                var setting = settingsPart.Settings.GetFirstChild<FootnoteDocumentWideProperties>() ?? settingsPart.Settings.AppendChild(new FootnoteDocumentWideProperties());
                setting.NumberingFormat = new NumberingFormat() { Val = format.Value };
                return true;
            }
        }
        else if (cw.Name?.StartsWith("aftnn") == true)
        {
            var format = RtfNumberFormatMapper.GetNumberFormat(cw.Name);
            if (format != null)
            {        
                settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                var setting = settingsPart.Settings.GetFirstChild<EndnoteDocumentWideProperties>() ?? settingsPart.Settings.AppendChild(new EndnoteDocumentWideProperties());
                setting.NumberingFormat = new NumberingFormat() { Val = format.Value };
                return true;
            }           
        }
        else if (cw.Name?.StartsWith("sftnn") == true)
        {
            var format = RtfNumberFormatMapper.GetNumberFormat(cw.Name);
            if (format != null)
            {
                currentSectPr ??= new SectionProperties();
                var footnotePr = currentSectPr.GetFirstChild<FootnoteProperties>() ?? currentSectPr.AppendChild(new FootnoteProperties());
                footnotePr.NumberingFormat = new NumberingFormat() { Val = format.Value };
                return true;
            }          
        }
        else if (cw.Name?.StartsWith("saftnn") == true)
        {
            var format = RtfNumberFormatMapper.GetNumberFormat(cw.Name);
            if (format != null)
            {
                currentSectPr ??= new SectionProperties();
                var endnotePr = currentSectPr.GetFirstChild<EndnoteProperties>() ?? currentSectPr.AppendChild(new EndnoteProperties());
                endnotePr.NumberingFormat = new NumberingFormat() { Val = format.Value };
                return true;
            }
        }

        return false;
    }

    private void ProcessFootnoteContinuationNotice(RtfGroup group)
    {
        ProcessPartContent(group, () => CreateSpecialFootnoteEndnote(FootnoteEndnoteValues.ContinuationNotice, false));
    }

    private void ProcessFootnoteContinuationSeparator(RtfGroup group)
    {
        ProcessPartContent(group, () => CreateSpecialFootnoteEndnote(FootnoteEndnoteValues.ContinuationSeparator, false));
    }

    private void ProcessFootnoteSeparator(RtfGroup group)
    {
        ProcessPartContent(group, () => CreateSpecialFootnoteEndnote(FootnoteEndnoteValues.Separator, false));
    }

    private void ProcessEndnoteContinuationNotice(RtfGroup group)
    {
        ProcessPartContent(group, () => CreateSpecialFootnoteEndnote(FootnoteEndnoteValues.ContinuationNotice, true));
    }

    private void ProcessEndnoteContinuationSeparator(RtfGroup group)
    {
        ProcessPartContent(group, () => CreateSpecialFootnoteEndnote(FootnoteEndnoteValues.ContinuationSeparator, true));
    }

    private void ProcessEndnoteSeparator(RtfGroup group)
    {
        ProcessPartContent(group, () => CreateSpecialFootnoteEndnote(FootnoteEndnoteValues.Separator, true));
    }
    
    private void ProcessFootnoteEndnote(RtfGroup group)
    {
        if (!pendingFootnoteEndnoteRef)
            return;
        
        // Determine if the group is a footnote or endnote
        bool isEndnote = group.Tokens.FirstOrDefault() is RtfControlWord cw && cw.Name.Equals("ftnalt", StringComparison.OrdinalIgnoreCase);
        FootnoteEndnoteReferenceType reference = isEndnote ? new EndnoteReference() : new FootnoteReference();        
        long id;
        if (isEndnote)
        {
            var endnotesPart = mainPart.EndnotesPart ?? mainPart.AddNewPart<EndnotesPart>();
            endnotesPart.Endnotes ??= new Endnotes();
            var ids = endnotesPart.Endnotes.OfType<Endnote>().Where(x => x.Id != null).Select(x => x.Id!.Value);
            if (ids.Count() == 0)
                id = 1;
            else 
                id = Math.Max(1, ids.Max());
        }
        else
        {
            var footnotesPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes ??= new Footnotes();
            var ids = footnotesPart.Footnotes.OfType<Footnote>().Where(x => x.Id != null).Select(x => x.Id!.Value);
            if (ids.Count() == 0)
                id = 1;
            else 
                id = Math.Max(1, ids.Max());
        }
        reference.Id = id;

        // Add foonote/endnote reference to a new run, ensuring proper formatting
        CreateRun().Append(isEndnote ? new EndnoteReference() { Id = id } : new FootnoteReference() { Id = id });
        pendingFootnoteEndnoteRef = false;
        
        // Add content to the footnote or endnote
        ProcessPartContent(group, () => CreateFootnoteEndnote(id, isEndnote));
    }

    private FootnoteEndnoteType CreateFootnoteEndnote(long id, bool isEndnote)
    {
        if (isEndnote)
        {
            var endnotesPart = mainPart.EndnotesPart ?? mainPart.AddNewPart<EndnotesPart>();
            endnotesPart.Endnotes ??= new Endnotes();
            return endnotesPart.Endnotes.AppendChild(new Endnote() { Id = id });
        }
        else
        {
            var footnotesPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes ??= new Footnotes();
            return footnotesPart.Footnotes.AppendChild(new Footnote() { Id = id });
        }
    }

    private FootnoteEndnoteType CreateSpecialFootnoteEndnote(FootnoteEndnoteValues type, bool isEndnote)
    {
        if (isEndnote)
        {
            var endnotesPart = mainPart.EndnotesPart ?? mainPart.AddNewPart<EndnotesPart>();
            endnotesPart.Endnotes ??= new Endnotes();
            return endnotesPart.Endnotes.AppendChild(new Endnote() { Type = type });
        }
        else
        {
            var footnotesPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes ??= new Footnotes();
            return footnotesPart.Footnotes.AppendChild(new Footnote() { Type = type });
        }
    }
}