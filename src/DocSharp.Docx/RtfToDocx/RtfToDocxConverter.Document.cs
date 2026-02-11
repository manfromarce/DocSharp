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
    private bool ProcessDocumentControlWord(RtfControlWord cw)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            // RTF header
            case "ansi":
                // If ANSI is specified, use the system ANSI code page, 
                // unless the DefaultCodePage value is set to a different value. 
                // Note that this default encoding can still be superseded by the \ansicpgN control word, if found. 
                int defaultCodePage;
                if (DefaultCodePage != null && DefaultCodePage.Value > 0)
                    defaultCodePage = DefaultCodePage.Value;
                else 
                    defaultCodePage = CultureInfo.CurrentCulture.TextInfo.ANSICodePage;

                codePageEncoding = Encoding.GetEncoding(defaultCodePage);
                return true;
            case "mac": // Legacy Mac encoding
                codePageEncoding = Encoding.GetEncoding(10000);
                // Note: 10000 is Mac Roman, but other encodings exist: 
                // MAC Japan (10001), MAC Arabic (10004), MAC Hebrew (10005), MAC Greek (10006), MAC Cyrillic (10007), MAC Latin2 (10029), MAC Turkish (10081)
                // For now, assume these would be specified in \ansicpg
                return true;
            case "pc": // IBM PC code page 437
                codePageEncoding = Encoding.GetEncoding(437);
                return true;
            case "pca": // BM PC code page 850
                codePageEncoding = Encoding.GetEncoding(850);
                return true;
            case "ansicpg": // If present, this control word should be after \ansi or \mac
                if (cw.HasValue && cw.Value!.Value >= 0)
                {
                    try
                    {
                        codePageEncoding = Encoding.GetEncoding(cw.Value.Value);                        
                    }
                    catch
                    {
#if DEBUG
                        Debug.WriteLine($"Unsupported code page: {cw.Value.Value}");
#endif                        
                    }
                }
                return true;

            // Document settings
            case "bookfold":
                CreateSetting<BookFoldPrinting>(true); 
                return true;
            case "bookfoldrev":
                CreateSetting<BookFoldReversePrinting>(true); 
                return true;
            // TODO: default fonts
            // case "deff":
            // case "adeff":
            // case "stshfdbch":
            // case "stshfloch":
            // case "stshfhich":
            // case "stshfbi":
            //     return true;
            case "deflang":
            case "deflangfe":
            case "adeflang":
                if (cw.HasValue && cw.Value != 1024 
                    && RtfHelpers.GetLanguageId(cw.Value!.Value) is string defaultLangId && !string.IsNullOrWhiteSpace(defaultLangId) 
                    && EnsureDocDefaults<Languages>() is Languages defaultLang)
                {
                    if (name == "deflang")
                        defaultLang.Val = defaultLangId;
                    if (name == "deflangfe")
                        defaultLang.EastAsia = defaultLangId;
                    if (name == "adeflang")
                        defaultLang.Bidi = defaultLangId;
                }
                return true;
            case "deftab":
                if (cw.HasValue && cw.Value > 0)
                {
                    CreateSetting<DefaultTabStop>((short)cw.Value!.Value);
                }
                return true;
            case "facingp":
                CreateSetting<EvenAndOddHeaders>(true);
                return true;
            case "formprot":
                defaultSectPr ??= new SectionProperties();
                var defaultProt = defaultSectPr.GetFirstChild<FormProtection>() ?? defaultSectPr.AppendChild(new FormProtection());
                defaultProt.Val = false;
                return true;
            case "formshade":
                CreateSetting<DoNotShadeFormData>(false); 
                return true;
            case "gutter":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Gutter = (uint)cw.Value!.Value;
                }
                return true;
            case "gutterprl": 
                CreateSetting<GutterAtTop>(true);
                return true;
            case "hyphauto":
                // \hyphauto and \hyphauto1 turn hyphenation on, \hyphauto0 turns hyphenation off
                CreateSetting<AutoHyphenation>(cw.Value, true); 
                return true;
            case "landscape":
                defaultSectPr ??= new SectionProperties();
                var defaultPgSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                defaultPgSize.Orient = PageOrientationValues.Landscape;
                return true;
            case "margb":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Bottom = cw.Value!.Value;
                }
                return true;
            case "margl":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Left = (uint)cw.Value!.Value;
                }
                return true;
            case "margr":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Right = (uint)cw.Value!.Value;
                }
                return true;
            case "margt":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageMargin = defaultSectPr.GetFirstChild<PageMargin>() ?? defaultSectPr.AppendChild(new PageMargin());
                    pageMargin.Top = cw.Value!.Value;
                }
                return true;
            case "margmirror": 
                CreateSetting<MirrorMargins>(true);
                return true;
            // case "ogutter": // Outside gutter, not used by Word (not sure how it should be mapped)
            //     return true;
            case "paperw":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                    pageSize.Width = (uint)cw.Value!.Value;
                }
                return true;
            case "paperh":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                    pageSize.Height = (uint)cw.Value!.Value;
                }
                return true;            
            case "pgbrdrhead":
                CreateSetting<BordersDoNotSurroundHeader>(false); 
                return true;
            case "pgbrdrfoot":
                CreateSetting<BordersDoNotSurroundFooter>(false); 
                return true;
            case "pgbrdrsnap":
                CreateSetting<AlignBorderAndEdges>(true); 
                return true;
            case "pgnstart":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageNumbers = defaultSectPr.GetFirstChild<PageNumberType>() ?? defaultSectPr.AppendChild(new PageNumberType());
                    pageNumbers.Start = cw.Value!.Value;
                }
                return true;
            case "printdata":
                CreateSetting<PrintFormsData>(true); 
                return true;
            case "psz":
                if (cw.HasValue)
                {
                    defaultSectPr ??= new SectionProperties();
                    var pageSize = defaultSectPr.GetFirstChild<PageSize>() ?? defaultSectPr.AppendChild(new PageSize());
                    pageSize.Code = (ushort)cw.Value!.Value;
                }
                return true;
            case "readonlyrecommended":
                settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                var setting = settingsPart.Settings.WriteProtection ?? settingsPart.Settings.AppendChild(new WriteProtection());
                setting.Recommended = true;
                return true;
            case "remdttm":
                CreateSetting<RemoveDateAndTime>(true); 
                return true;
            case "rempersonalinfo":
                CreateSetting<RemovePersonalInformation>(true); 
                return true;
            case "showplaceholdtext":
                CreateSetting<AlwaysShowPlaceholderText>(cw.Value, null); 
                return true;
            case "twoonone":
                CreateSetting<PrintTwoOnOne>(cw.Value, null); 
                return true;
            case "viewkind":
                if (cw.HasValue)
                {
                    if (cw.Value == 0)
                    {
                        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings ??= new Settings();
                        var view = settingsPart.Settings.View ?? settingsPart.Settings.AppendChild(new View());
                        view.Val = ViewValues.None;
                    }
                    else if (cw.Value == 1)
                    {
                        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings ??= new Settings();
                        var view = settingsPart.Settings.View ?? settingsPart.Settings.AppendChild(new View());
                        view.Val = ViewValues.Print;
                    }
                    else if (cw.Value == 2)
                    {
                        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings ??= new Settings();
                        var view = settingsPart.Settings.View ?? settingsPart.Settings.AppendChild(new View());
                        view.Val = ViewValues.Outline;
                    }
                    else if (cw.Value == 3)
                    {
                        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings ??= new Settings();
                        var view = settingsPart.Settings.View ?? settingsPart.Settings.AppendChild(new View());
                        view.Val = ViewValues.MasterPages;
                    }
                    else if (cw.Value == 4)
                    {
                        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings ??= new Settings();
                        var view = settingsPart.Settings.View ?? settingsPart.Settings.AppendChild(new View());
                        view.Val = ViewValues.Normal;
                    }
                    else if (cw.Value == 5)
                    {
                        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings ??= new Settings();
                        var view = settingsPart.Settings.View ?? settingsPart.Settings.AppendChild(new View());
                        view.Val = ViewValues.Web;
                    }
                }
                return true;
        }
        return false;
    }

    private T EnsureDocDefaults<T>() where T: OpenXmlElement, new()
    {
        stylesPart ??= mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();
        var docDefaults = stylesPart.Styles.DocDefaults ?? stylesPart.Styles.AppendChild(new DocDefaults());
        return docDefaults.GetFirstChild<T>() ?? docDefaults.AppendChild(new T());
    }

    private void CreateSetting<T>(short rtfValue) where T: NonNegativeShortType, new()
    {
        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        var setting = settingsPart.Settings.GetFirstChild<T>() ?? settingsPart.Settings.AppendChild(new T());
        setting.Val = rtfValue;
    }

    private void CreateSetting<T>(int? rtfValue, bool? defaultIfNoValue) where T: OnOffType, new()
    {
        if (rtfValue == null && defaultIfNoValue == null)
            return;
        bool value = (rtfValue != null) ? (rtfValue.Value != 0) : defaultIfNoValue!.Value;
        CreateSetting<T>(value);
    }

    private void CreateSetting<T>(bool value) where T: OnOffType, new()
    {
        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        var setting = settingsPart.Settings.GetFirstChild<T>() ?? settingsPart.Settings.AppendChild(new T());
        setting.Val = value;
    }

    private T EnsureSetting<T>() where T : OpenXmlElement, new()
    {
        settingsPart ??= mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        return settingsPart.Settings.GetFirstChild<T>() ?? settingsPart.Settings.AppendChild(new T());
    }
}