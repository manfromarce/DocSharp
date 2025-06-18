using System.IO;
using System.Xml;
using System.Linq;
using System.Text;
using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Rtf.Tokens;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Globalization;
using System.Web;
using System.Net;
using DocSharp.Helpers;

namespace DocSharp.Rtf
{
    internal class DocxBuilder
    {
        private WordprocessingDocument _doc;
        private MainDocumentPart _mainPart;
        private DocumentFormat.OpenXml.Wordprocessing.Document _document;
        private Body _body;
        private SectionProperties? _currentSectionProperties;
        private SectionProperties? _defaultSectionProperties;
        private ParagraphProperties? _currentParagraphProperties;
        private RunProperties? _currentRunProperties;
        private Stack<RunProperties> _formattingStack = new Stack<RunProperties>();
        private Paragraph? _currentParagraph;
        private Run? _currentRun;

        public DocxBuilder(WordprocessingDocument doc)
        {
            _doc = doc;
            _mainPart = doc.AddMainDocumentPart();
            _document = _mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            _body = _mainPart.Document.AppendChild(new Body());
            _defaultSectionProperties = CreateDefaultSectionProperties();
            _currentSectionProperties = _defaultSectionProperties;
        }

        /// <summary>
        /// Convert RTF to Open XML and add to the document.
        /// </summary>
        /// <param name="document">The parsed RTF document</param>
        public void AddRtf(RtfDocument document)
        {
            foreach (var token in document.Contents)
            {
                if (token is FontTableTag || token is ColorTable)
                {
                    break;
                }
                if (token.Type == TokenType.Group)
                {
                    ProcessGroup((Group)token);
                }
                else if (token.Type == TokenType.BreakTag)
                {
                    ProcessBreak(token);
                }
                else if (token.Type == TokenType.Text)
                {
                    ProcessText((TextToken)token);
                }
                else
                {
                    ProcessDocumentFormatting(token);
                    ProcessSectionFormatting(token);
                    ProcessParagraphFormatting(token);
                    ProcessCharacterFormatting(token);
                }
            }
        }

        private void ProcessBreak(IToken token)
        {
            if (token is IWord word)
            {
                if (word.Name == "pagebb" || word.Name == "page")
                {
                    EnsureCurrentRun();
                    _currentRun!.AppendChild(new Break() { Type = BreakValues.Page });
                }
                else if (word.Name == "line")
                {
                    EnsureCurrentRun();
                    _currentRun!.AppendChild(new Break() { Type = BreakValues.TextWrapping });
                }
                else if (word.Name == "column")
                {
                    EnsureCurrentRun();
                    _currentRun!.AppendChild(new Break() { Type = BreakValues.Column });
                }
                else if (word.Name == "par")
                {
                    _currentParagraph = null;
                    EnsureCurrentParagraph();
                }
                else if (word.Name == "sect")
                {
                    _currentSectionProperties = CreateSectionProperties();
                }
            }
        }

        private void ProcessText(TextToken textToken)
        {
            EnsureCurrentParagraph();

            var run = new Run();

            if (_currentRunProperties != null)
            {
                run.AppendChild(_currentRunProperties.CloneNode(true));
            }

            string escapedText = WebUtility.HtmlEncode(textToken.Value);

            run.AppendChild(new Text(escapedText) { Space = SpaceProcessingModeValues.Preserve });

            _currentParagraph?.AppendChild(run);
            _currentRun = run;
        }

        private void ProcessGroup(Group group)
        {
            if (group.Destination == null)
            {
                ProcessGroupContent(group);
            }
            else
            {
                switch (group.Destination.Name)
                {
                    case "aftncn":
                        break;

                    case "aftnsep":
                        break;
                    case "aftnsepc":
                        break;
                    case "annotation":
                        break;
                    case "atnauthor":
                        break;
                    case "atndate":
                        break;
                    case "atnicn":
                        break;
                    case "atnid":
                        break;
                    case "atnparent":
                        break;
                    case "atnref":
                        break;
                    case "atntime":
                        break;
                    case "atrfend":
                        break;
                    case "atrfstart":
                        break;
                    case "author":
                        break;

                    case "background":
                        break;
                    case "bkmkend":
                        break;
                    case "bkmkstart":
                        break;
                    case "blipuid":
                        break;

                    case "buptim":
                        break;

                    case "category":
                        break;

                    case "colorschememapping":
                        break;
                    case "colortbl":
                        break;
                    case "comment":
                        break;
                    case "company":
                        break;
                    case "creatim":
                        break;
                    case "datastore":
                        break;
                    case "datafield":
                        break;
                    case "defchp":
                        break;

                    case "defpap":
                        break;
                    case "do":
                        break;
                    case "doccomm":
                        break;
                    case "docvar":
                        break;

                    case "dptxbxtext":
                        break;

                    case "ebcend":
                        break;
                    case "ebcstart":
                        break;
                    case "factoidname":
                        break;
                    case "falt":
                        break;

                    case "fchars":
                        break;
                    case "ffdeftext":
                        break;
                    case "ffentrymcr":
                        break;
                    case "ffexitmcr":
                        break;
                    case "ffformat":
                        break;
                    case "ffhelptext":
                        break;
                    case "ffl":
                        break;
                    case "ffname":
                        break;
                    case "ffstattext":
                        break;
                    case "field":
                        break;
                    case "file":
                        break;
                    case "filetbl":
                        break;

                    case "fldinst":
                        break;
                    case "fldrslt":
                        break;
                    //case "fldtype":
                    //    break;
                    case "fname":
                        break;
                    case "fontemb":
                        break;
                    case "fontfile":
                        break;
                    case "fonttbl":
                        break;
                    case "footer":
                        break;
                    case "footerf":
                        break;
                    case "footerl":
                        break;
                    case "footerr":
                        break;
                    case "footnote":
                        break;
                    case "formfield":
                        break;

                    case "ftncn":
                        break;
                    case "ftnsep":
                        break;
                    case "ftnsepc":
                        break;

                    case "g":
                        break;
                    case "generator":
                        break;
                    case "gridtbl":
                        break;
                    case "header":
                        break;
                    case "headerf":
                        break;
                    case "headerl":
                        break;
                    case "headerr":
                        break;
                    case "hl":
                        break;
                    case "hlfr":
                        break;
                    case "hlinkbase":
                        break;
                    case "hlloc":
                        break;
                    case "hlsrc":
                        break;
                    case "hsv":
                        break;
                    case "htmltag":
                        break;

                    case "info":
                        break;
                    case "keycode":
                        break;
                    case "keywords":
                        break;

                    case "latentstyles":
                        break;
                    case "lchars":
                        break;
                    case "levelnumbers":
                    case "leveltext":
                    case "lfolevel":
                        break;
                    case "linkval":
                        break;
                    case "list":
                    case "listlevel":
                    case "listname":
                    case "listoverride":
                    case "listoverridetable":
                    case "listpicture":
                    case "liststylename":
                    case "listtable":
                    case "listtext":
                        break;
                    case "lsdlockedexcept":
                        break;
                    case "macc":
                        break;
                    case "maccPr":
                        break;
                    case "mailmerge":
                        break;
                    case "maln":
                        break;
                    case "malnScr":
                        break;
                    case "manager":
                        break;
                    case "margPr":
                        break;
                    case "mbar":
                        break;
                    case "mbarPr":
                        break;
                    case "mbaseJc":
                        break;
                    case "mbegChr":
                        break;
                    case "mborderBox":
                        break;
                    case "mborderBoxPr":
                        break;
                    case "mbox":
                        break;
                    case "mboxPr":
                        break;
                    case "mchr":
                        break;
                    case "mcount":
                        break;
                    case "mctrlPr":
                        break;
                    case "md":
                        break;
                    case "mdeg":
                        break;
                    case "mdegHide":
                        break;
                    case "mden":
                        break;
                    case "mdiff":
                        break;
                    case "mdPr":
                        break;
                    case "me":
                        break;
                    case "mendChr":
                        break;
                    case "meqArr":
                        break;
                    case "meqArrPr":
                        break;
                    case "mf":
                        break;
                    case "mfName":
                        break;
                    case "mfPr":
                        break;
                    case "mfunc":
                        break;
                    case "mfuncPr":
                        break;
                    case "mgroupChr":
                        break;
                    case "mgroupChrPr":
                        break;
                    case "mgrow":
                        break;
                    case "mhideBot":
                        break;
                    case "mhideLeft":
                        break;
                    case "mhideRight":
                        break;
                    case "mhideTop":
                        break;
                    case "mhtmltag":
                        break;
                    case "mlim":
                        break;
                    case "mlimloc":
                        break;
                    case "mlimlow":
                        break;
                    case "mlimlowPr":
                        break;
                    case "mlimupp":
                        break;
                    case "mlimuppPr":
                        break;
                    case "mm":
                        break;
                    case "mmaddfieldname":
                        break;
                    case "mmath":
                        break;
                    case "mmathPict":
                        break;
                    case "mmathPr":
                        break;
                    case "mmaxdist":
                        break;
                    case "mmc":
                        break;
                    case "mmcJc":
                        break;
                    case "mmconnectstr":
                        break;
                    case "mmconnectstrdata":
                        break;
                    case "mmcPr":
                        break;
                    case "mmcs":
                        break;
                    case "mmdatasource":
                    case "mmheadersource":
                    case "mmmailsubject":
                    case "mmodso":
                    case "mmodsomappedname":
                    case "mmodsofilter":
                    case "mmodsofldmpdata":
                    case "mmodsoname":
                    case "mmodsorecipdata":
                    case "mmodsosort":
                    case "mmodsosrc":
                    case "mmodsotable":
                    case "mmodsoudl":
                    case "mmodsoudldata":
                    case "mmodsouniquetag":
                        break;
                    case "mmPr":
                        break;
                    case "mmquery":
                        break;
                    case "mmr":
                        break;
                    case "mnary":
                        break;
                    case "mnaryPr":
                        break;
                    case "mnoBreak":
                        break;
                    case "mnum":
                        break;
                    case "mobjDist":
                        break;
                    case "moMath":
                        break;
                    case "moMathPara":
                        break;
                    case "moMathParaPr":
                        break;
                    case "mopEmu":
                        break;
                    case "mphant":
                        break;
                    case "mphantPr":
                        break;
                    case "mplcHide":
                        break;
                    case "mpos":
                        break;
                    case "mr":
                        break;
                    case "mrad":
                        break;
                    case "mradPr":
                        break;
                    case "mrPr":
                        break;
                    case "msepChr":
                        break;
                    case "mshow":
                        break;
                    case "mshp":
                        break;
                    case "msPre":
                        break;
                    case "msPrePr":
                        break;
                    case "msSub":
                        break;
                    case "msSubPr":
                        break;
                    case "msSubSup":
                        break;
                    case "msSubSupPr":
                        break;
                    case "msSup":
                        break;
                    case "msSupPr":
                        break;
                    case "mstrikeBLTR":
                        break;
                    case "mstrikeH":
                        break;
                    case "mstrikeTLBR":
                        break;
                    case "mstrikeV":
                        break;
                    case "msub":
                        break;
                    case "msubHide":
                        break;
                    case "msup":
                        break;
                    case "msupHide":
                        break;
                    case "mtransp":
                        break;
                    case "mtype":
                        break;
                    case "mvertJc":
                        break;
                    case "mvfmf":
                    case "mvfml":
                    case "mvtof":
                    case "mvtol":
                        break;
                    case "mzeroAsc":
                        break;
                    case "mzeroDesc":
                        break;
                    case "mzeroWid":
                        break;
                    case "nesttableprops":
                        break;
                    case "nextfile":
                        break;
                    case "nonesttables":
                        break;
                    case "nonshppict":
                        break;
                    case "objalias":
                        break;
                    case "objclass":
                        break;
                    case "object":
                        break;
                    case "objdata":
                        break;
                    case "objname":
                        break;
                    case "objsect":
                        break;
                    case "objtime":
                        break;
                    case "oldcprops":
                        break;
                    case "oldpprops":
                        break;
                    case "oldsprops":
                        break;
                    case "oldtprops":
                        break;
                    case "oleclsid":
                        break;
                    case "operator":
                        break;
                    case "panose":
                        break;
                    case "password":
                        break;
                    case "passwordhash":
                        break;
                    case "pgp":
                        break;
                    case "pgptbl":
                        break;
                    case "picprop":
                        break;
                    case "pict":
                        break;
                    case "pn":
                        break;
                    case "pntext":
                        break;
                    case "pntxta":
                        break;
                    case "pntxtb":
                        break;
                    case "printim":
                        break;
                    case "private":
                        break;
                    case "propname":
                        break;
                    case "protend":
                        break;
                    case "protstart":
                        break;
                    case "protusertbl":
                        break;
                    case "pxe":
                        break;
                    case "result":
                        break;
                    case "revtbl":
                        break;
                    case "revtim":
                        break;
                    case "rsidtbl":
                        break;
                    case "rtf":
                        break;
                    case "rxe":
                        break;
                    case "shp":
                        break;
                    case "shpgrp":
                        break;
                    case "shpinst":
                        break;
                    case "shppict":
                        break;
                    case "shprslt":
                        break;
                    case "shptxt":
                        break;
                    case "sn":
                        break;
                    case "sp":
                        break;
                    case "staticval":
                        break;
                    case "stylesheet":
                        break;
                    case "subject":
                        break;
                    case "sv":
                        break;
                    case "svb":
                        break;
                    case "tc":
                        break;
                    case "template":
                        break;
                    case "themedata":
                        break;
                    case "title":
                        break;
                    case "txe":
                        break;
                    case "ud":
                        break;
                    case "upr":
                        break;
                    case "userprops":
                        break;
                    case "wgrffmtfilter":
                        break;
                    case "windowcaption":
                        break;
                    case "writereservation":
                        break;
                    case "writereservhash":
                        break;
                    case "xe":
                        break;
                    case "xform":
                        break;
                    case "xmlattrname":
                        break;
                    case "xmlattrvalue":
                        break;
                    case "xmlclose":
                        break;
                    case "xmlname":
                        break;
                    case "xmlnstbl":
                        break;
                    case "xmlopen":
                        break;
                    case "fldtype":
                        break;
                    case "*":
                        break;
                    default:
                        if (group.Destination.Name.StartsWith("htmltag", StringComparison.OrdinalIgnoreCase) ||
                            group.Destination.Name.StartsWith("mhtmltag", StringComparison.OrdinalIgnoreCase) ||
                            group.Destination.Name.StartsWith("pnseclvl", StringComparison.OrdinalIgnoreCase) ||
                            group.Destination.Name.StartsWith("ebcstart", StringComparison.OrdinalIgnoreCase) ||
                            group.Destination.Name.StartsWith("ebcend", StringComparison.OrdinalIgnoreCase))
                        {

                        }
                        else if (group.Contents.Count > 0 && group.Contents[0] is IgnoreUnrecognized)
                        {

                        }
                        else
                        {
                            ProcessGroupContent(group);
                        }
                        break;
                }
            }
        }

        internal void ProcessGroupContent(Group group)
        {
            var newRunProperties = _currentRunProperties?.CloneNode(true) as RunProperties ?? new RunProperties();
            // TODO: paragraph properties?
            _formattingStack.Push(newRunProperties);

            // Process child tokens
            foreach (var token in group.Contents)
            {
                if (token.Type == TokenType.Group)
                {
                    ProcessGroup((Group)token);
                }
                else if (token.Type == TokenType.BreakTag)
                {
                    ProcessBreak(token);
                }
                else if (token.Type == TokenType.Text)
                {
                    ProcessText((TextToken)token);
                }
                else
                {
                    ProcessDocumentFormatting(token);
                    ProcessSectionFormatting(token);
                    ProcessParagraphFormatting(token);
                    ProcessCharacterFormatting(token);
                }
            }

            _formattingStack.Pop();
            _currentRunProperties = _formattingStack.Any() ? _formattingStack.Peek() : null;
        }

        internal void ProcessDocumentFormatting(IToken token)
        {
            if (token is IWord word)
            {
                switch (word.Name)
                {
                    case "rtlgutter":
                        _defaultSectionProperties ??= CreateDefaultSectionProperties();
                        if (!_defaultSectionProperties.Elements<GutterOnRight>().Any())
                        {
                            _defaultSectionProperties.AppendChild(new GutterOnRight());
                        }
                        return;
                }
            }
            ProcessPageInformation(token);
        }

        internal void ProcessSectionFormatting(IToken token)
        {
            if (token is IWord word)
            {
                switch (word.Name)
                {
                    case "sectd":
                        if (_defaultSectionProperties != null)
                        {
                            _currentSectionProperties = (SectionProperties)_defaultSectionProperties.CloneNode(true);
                        }
                        else
                        {
                            // Should not happen
                            _currentSectionProperties = CreateDefaultSectionProperties();
                        }
                        return;
                    case "sbkpage":
                        _currentSectionProperties ??= CreateSectionProperties();
                        SetSectionType(_currentSectionProperties, SectionMarkValues.NextPage);
                        return;
                    case "sbkcol":
                        _currentSectionProperties ??= CreateSectionProperties();
                        SetSectionType(_currentSectionProperties, SectionMarkValues.NextColumn);
                        return;
                    case "sbkcont":
                        _currentSectionProperties ??= CreateSectionProperties();
                        SetSectionType(_currentSectionProperties, SectionMarkValues.Continuous);
                        return;
                    case "sbkodd":
                        _currentSectionProperties ??= CreateSectionProperties();
                        SetSectionType(_currentSectionProperties, SectionMarkValues.OddPage);
                        return;
                    case "sbkeven":
                        _currentSectionProperties ??= CreateSectionProperties();
                        SetSectionType(_currentSectionProperties, SectionMarkValues.EvenPage);
                        return;
                    case "titlepg":
                        _currentSectionProperties ??= CreateSectionProperties();
                        if (!_currentSectionProperties.Elements<TitlePage>().Any())
                        {
                            _currentSectionProperties.AppendChild(new TitlePage());
                        }
                        return;
                }
            }
            ProcessPageInformation(token);
        }

        internal void ProcessParagraphFormatting(IToken token)
        {
            if (token is IWord word)
            {
                switch (word.Name)
                {
                    case "pard":
                        ResetParagraphProperties();
                        return;

                    case "li":
                        SetParagraphProperty<Indentation>(c => c.Left = ((ControlWord<UnitValue>)word).Value.ToTwip().ToStringInvariant());
                        return;
                    case "ri":
                        SetParagraphProperty<Indentation>(c => c.Right = ((ControlWord<UnitValue>)word).Value.ToTwip().ToStringInvariant());
                        return;
                    case "fi":
                        int firstLineIndent = ((ControlWord<UnitValue>)word).Value.ToTwip();
                        if (firstLineIndent >= 0)
                            SetParagraphProperty<Indentation>(c => c.FirstLine = firstLineIndent.ToStringInvariant());
                        else 
                            SetParagraphProperty<Indentation>(c => c.Hanging = Math.Abs(firstLineIndent).ToStringInvariant());
                        return;
                    case "culi":
                        SetParagraphProperty<Indentation>(c => c.LeftChars = ((ControlWord<int>)word).Value);
                        return;
                    case "curi":
                        SetParagraphProperty<Indentation>(c => c.RightChars = ((ControlWord<int>)word).Value);
                        return;
                    case "cufi":
                        int firstLineIndentChars = ((ControlWord<int>)word).Value;
                        if (firstLineIndentChars >= 0)
                            SetParagraphProperty<Indentation>(c => c.FirstLineChars = firstLineIndentChars);
                        else
                            SetParagraphProperty<Indentation>(c => c.HangingChars = Math.Abs(firstLineIndentChars));
                        return;
                    case "contextualspace":
                        SetParagraphProperty<ContextualSpacing>(c => c.Val = true);
                        return;
                    case "ql":
                        SetParagraphProperty<Justification>(c => c.Val = JustificationValues.Left);
                        return;
                    case "qc":
                        SetParagraphProperty<Justification>(c => c.Val = JustificationValues.Center);
                        return;
                    case "qr":
                        SetParagraphProperty<Justification>(c => c.Val = JustificationValues.Right);
                        return;
                    case "qj":
                        SetParagraphProperty<Justification>(c => c.Val = JustificationValues.Both);
                        return;
                    case "qd":
                        SetParagraphProperty<Justification>(c => c.Val = JustificationValues.Distribute);
                        return;
                    case "qt":
                        SetParagraphProperty<Justification>(c => c.Val = JustificationValues.ThaiDistribute);
                        return;
                    case "qk":
                        int qkVal = ((ControlWord<int>)word).Value;
                        if (qkVal == 0)
                            SetParagraphProperty<Justification>(c => c.Val = JustificationValues.LowKashida);
                        else if (qkVal == 10)
                            SetParagraphProperty<Justification>(c => c.Val = JustificationValues.MediumKashida);
                        else if (qkVal == 20)
                            SetParagraphProperty<Justification>(c => c.Val = JustificationValues.HighKashida);
                        return;
                        //case "sa":
                        //case "saauto":
                        //case "sb":
                        //case "sbauto":
                        //case "lisa":
                        //case "lisb":
                        //case "sl":
                        //case "slmult":
                        //case "adjustright":
                        //case "nosnaplinegrid":
                        //    return;
                }
            }
        }

        private void ResetParagraphProperties()
        {
            EnsureCurrentParagraph();
            if (_currentParagraph!.GetFirstChild<ParagraphProperties>() is ParagraphProperties pPr)
            {
                _currentParagraph.RemoveChild(pPr);
            }
            _currentParagraphProperties = CreateDefaultParagraphProperties();
            _currentParagraph!.PrependChild(_currentParagraphProperties);
        }

        internal void ProcessCharacterFormatting(IToken token)
        {
            if (token is Font font)
            {
                //switch (font.Category)
                //{
                //    case FontFamilyCategory.Nil:
                        SetRunProperty<RunFonts>(c => c.Ascii = font.Name);
                        SetRunProperty<RunFonts>(c => c.ComplexScript = font.Name);
                        SetRunProperty<RunFonts>(c => c.EastAsia = font.Name);
                        SetRunProperty<RunFonts>(c => c.HighAnsi = font.Name);
                //        break;
                //}
            }
            else if (token is PositionOffset offset)
            {
                if (offset.Name == "up")
                {
                    SetRunProperty<Position>(c => c.Val = offset.Value.ToHalfPoints().ToStringInvariant());
                }
                else if (offset.Name == "dn")
                {
                    SetRunProperty<Position>(c => c.Val = "-" + offset.Value.ToHalfPoints().ToStringInvariant());
                }
            }
            else if (token is IWord word)
            {
                switch (word.Name)
                {
                    case "b": // Bold
                        SetRunProperty<Bold>(bold => bold.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "i": // Italic
                        SetRunProperty<Italic>(italic => italic.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "strike": 
                        SetRunProperty<Strike>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "striked": // Double strike
                        SetRunProperty<DoubleStrike>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "ul": // Single solid underline
                        SetRunProperty<Underline>(u => u.Val = ((ControlWord<bool>)word).Value ? UnderlineValues.Single : UnderlineValues.None);
                        return;
                    case "uld": // Dotted underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.Dotted);
                        return;
                    case "uldash": // Dashed underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.Dash);
                        return;
                    case "uldashd": // Dash-dot underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DotDash);
                        return;
                    case "uldashdd": // Dash-dot-dot underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DotDotDash);
                        return;
                    case "ulldash": // Long dash underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DashLong);
                        return;
                    case "ulthldash": // Thick long dash underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DashLongHeavy);
                        return;
                    case "ulth": // Thick underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.Thick);
                        return;
                    case "ulthd": // Thick dotted underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DottedHeavy);
                        return;
                    case "ulthdash": // Thick dashed underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DashedHeavy);
                        return;
                    case "ulthdashd": // Thick dash-dot underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DashDotHeavy);
                        return;
                    case "ulthdashdd": // Thick dash-dot-dot underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.DashDotDotHeavy);
                        return;
                    case "uldb": // Double underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.Double);
                        return;
                    case "ulwave": // Wavy underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.Wave);
                        return;
                    case "ululdbwave": // Double wavy underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.WavyDouble);
                        return;
                    case "ulhwave": // Thick wavy underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.WavyHeavy);
                        return;
                    case "ulw": // Words underline
                        SetRunProperty<Underline>(u => u.Val = UnderlineValues.Words);
                        return;
                    case "ulc": // Underline color
                        SetRunProperty<Underline>(u => u.Color = ((ControlWord<ColorValue>)word).Value.Red.ToString("X2") + ((ControlWord<ColorValue>)word).Value.Green.ToString("X2") + ((ControlWord<ColorValue>)word).Value.Blue.ToString("X2"));
                        return;
                    case "outl": // Outline
                        SetRunProperty<Outline>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "shad": // Shadow
                        SetRunProperty<Shadow>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "embo": // Embossed
                        SetRunProperty<Shadow>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "impr": // Engraved
                        SetRunProperty<Imprint>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "v": // Hidden
                        SetRunProperty<Hidden>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "scaps": // Small caps
                        SetRunProperty<SmallCaps>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "caps": // All caps
                        SetRunProperty<Caps>(s => s.Val = OnOffValue.FromBoolean(((ControlWord<bool>)word).Value));
                        return;
                    case "super": // Superscript
                        SetRunProperty<VerticalTextAlignment>(s => s.Val = VerticalPositionValues.Superscript);
                        return;
                    case "sub": // Subscript
                        SetRunProperty<VerticalTextAlignment>(s => s.Val = VerticalPositionValues.Subscript);
                        return;
                    case "nosupersub": // Disable superscript/subscript
                        SetRunProperty<VerticalTextAlignment>(s => s.Val = VerticalPositionValues.Baseline);
                        return;
                    //case "f": // FontRef (only if font could not be determined previously)
                    //    return;
                    case "fs": // Font size
                        SetRunProperty<W.FontSize>(c => c.Val = (((ControlWord<UnitValue>)word).Value.ToPt() * 2).ToStringInvariant());
                        return;                   
                    case "cf": // Font color
                        SetRunProperty<Color>(c => c.Val = ((ControlWord<ColorValue>)word).Value.Red.ToString("X2", CultureInfo.InvariantCulture) + ((ControlWord<ColorValue>)word).Value.Green.ToString("X2", CultureInfo.InvariantCulture) + ((ControlWord<ColorValue>)word).Value.Blue.ToString("X2", CultureInfo.InvariantCulture));
                        return;
                    //case "cb": // Background color (Word has never supported this control word and uses chcbpat or highlight instead)
                    //    return;
                    case "highlight": // Highlight color
                        EnumValue<HighlightColorValues>? highlight = ((ControlWord<ColorValue>)word).Value.ToHighlight();
                        if (highlight != null)
                            SetRunProperty<Highlight>(c => c.Val = highlight);
                        return;
                    case "chcbpat": // Character background color
                        SetRunProperty<Shading>(c => c.Fill = ((ControlWord<ColorValue>)word).Value.Red.ToString("X2", CultureInfo.InvariantCulture) + ((ControlWord<ColorValue>)word).Value.Green.ToString("X2", CultureInfo.InvariantCulture) + ((ControlWord<ColorValue>)word).Value.Blue.ToString("X2", CultureInfo.InvariantCulture));
                        return;
                    case "chcfpat": // Character foreground color
                        SetRunProperty<Shading>(c => c.Color = ((ControlWord<ColorValue>)word).Value.Red.ToString("X2", CultureInfo.InvariantCulture) + ((ControlWord<ColorValue>)word).Value.Green.ToString("X2", CultureInfo.InvariantCulture) + ((ControlWord<ColorValue>)word).Value.Blue.ToString("X2", CultureInfo.InvariantCulture));
                        return;
                    case "chshdng":
                        EnumValue<ShadingPatternValues>? shadingPattern = ((ControlWord<int>)word).ToShadingPattern();
                        if (shadingPattern != null)
                            SetRunProperty<Shading>(c => c.Val = shadingPattern);
                        return;
                    case "chbgdkcross":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.HorizontalCross);
                        return;
                    case "chbgcross":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ThinHorizontalCross);
                        return;
                    case "chbgdkhoriz":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.HorizontalStripe);
                        return;
                    case "chbghoriz":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ThinHorizontalStripe);
                        return;
                    case "chbgdkvert":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.VerticalStripe);
                        return;
                    case "chbgvert":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ThinVerticalStripe);
                        return;
                    case "chbgdkdcross":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.DiagonalCross);
                        return;
                    case "chbgdcross":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ThinDiagonalCross);
                        return;
                    case "chbgdkbdiag":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.DiagonalStripe);
                        return;
                    case "chbgbdiag":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ThinDiagonalStripe);
                        return;
                    case "chbgdkfdiag":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ReverseDiagonalStripe);
                        return;
                    case "chbgfdiag":
                        SetRunProperty<Shading>(c => c.Val = ShadingPatternValues.ThinReverseDiagonalStripe);
                        return;
                    case "charscalex": // Font scaling
                        SetRunProperty<CharacterScale>(c => c.Val = ((ControlWord<int>)word).Value);
                        return;
                    case "kerning": // Characters kerning
                        SetRunProperty<Kern>(c => c.Val = (uint)((ControlWord<int>)word).Value);
                        return;
                    case "fittext": // Fit text
                        SetRunProperty<W.FitText>(c => c.Val = (uint)((ControlWord<int>)word).Value);
                        return;
                    case "expnd": // Font spacing in quarter-points; convert to twips
                        SetRunProperty<Spacing>(c => c.Val = ((ControlWord<int>)word).Value / 5);
                        return;
                    case "expndtw": // Font spacing in twips
                        SetRunProperty<Spacing>(c => c.Val = ((ControlWord<int>)word).Value);
                        return;
                    case "noproof":
                        SetRunProperty<NoProof>(c => c.Val = OnOffValue.FromBoolean(true));
                        return;
                    case "accnone":
                        SetRunProperty<Emphasis>(c => c.Val = EmphasisMarkValues.None);
                        return;
                    case "accdot":
                        SetRunProperty<Emphasis>(c => c.Val = EmphasisMarkValues.Dot);
                        return;
                    case "acccomma":
                        SetRunProperty<Emphasis>(c => c.Val = EmphasisMarkValues.Comma);
                        return;
                    case "acccircle":
                        SetRunProperty<Emphasis>(c => c.Val = EmphasisMarkValues.Circle);
                        return;
                    case "accunderdot":
                        SetRunProperty<Emphasis>(c => c.Val = EmphasisMarkValues.UnderDot);
                        return;
                    //case "animtext" // Animation (not supported by Word 2007 and newer)
                    //case "gridtbl": // Destination keyword related to character grids (not emitted by Word)

                    // TODO
                    //case "chbrdr":
                    //case "cchs":
                    //case "cgrid":
                    //case "gcw":
                    //case "lang":
                    //case "langnp":
                    //case "langfe":
                    //case "langfenp":
                    //case "ltrch":
                    //case "rtlch":
                    //case "nosectexpand":
                    //case "webhidden":
                        //return;
                }
            }
        }

        private void SetRunProperty<T>(Action<T> setProperty) where T : OpenXmlElement, new()
        {
            _currentRunProperties ??= new RunProperties();
            var element = _currentRunProperties.Elements<T>().FirstOrDefault();
            if (element == null)
            {
                element = new T();
                _currentRunProperties.AppendChild(element);
            }
            setProperty(element);
        }

        internal void ProcessPageInformation(IToken token)
        {
            if (token is ControlWord<int> intWord)
            {
                switch (intWord.Name)
                {
                    // These apply to the document by default
                    case "paperw":
                    case "paperh":
                        bool isWidth = intWord.Name.Equals("paperw", StringComparison.OrdinalIgnoreCase);
                        _defaultSectionProperties ??= CreateDefaultSectionProperties();
                        SetPageDimensions(_defaultSectionProperties, intWord.Value, isWidth);
                        return;
                    case "margl":
                    case "margr":
                    case "margt":
                    case "margb":
                        _defaultSectionProperties ??= CreateDefaultSectionProperties();
                        string marginType = intWord.Name switch
                        {
                            "margl" => "Left",
                            "margr" => "Right",
                            "margt" => "Top",
                            "margb" => "Bottom",
                            "gutter" => "Gutter",
                            _ => throw new InvalidOperationException()
                        };
                        SetPageMargin(_defaultSectionProperties, intWord.Value, marginType);
                        return;

                    // These apply to the current section only
                    case "pgwsxn":
                    case "pghsxn":
                        bool isSectionWidth = intWord.Name.Equals("pgwsxn", StringComparison.OrdinalIgnoreCase);
                        _currentSectionProperties ??= CreateSectionProperties();
                        SetPageDimensions(_currentSectionProperties, intWord.Value, isSectionWidth);
                        return;
                    case "marglsxn":
                    case "margrsxn":
                    case "margtsxn":
                    case "margbsxn":
                    case "guttersxn":
                        _currentSectionProperties ??= CreateSectionProperties();
                        string sectionMarginType = intWord.Name switch
                        {
                            "marglsxn" => "Left",
                            "margrsxn" => "Right",
                            "margtsxn" => "Top",
                            "margbsxn" => "Bottom",
                            "guttersxn" => "Gutter",
                            "headery" => "Header",
                            "footery" => "Footer",
                            _ => throw new InvalidOperationException()
                        };
                        SetPageMargin(_currentSectionProperties, intWord.Value, sectionMarginType);
                        return;
                }
            }
        }

        internal SectionProperties CreateSectionProperties()
        {
            var sectionProperties = new SectionProperties();

            // Copy from previous section properties if present (RTF behavior, unless \sectd is specified)
            if (_currentSectionProperties != null)
            {
                foreach (var element in _currentSectionProperties.Elements())
                {
                    sectionProperties.AppendChild(element.CloneNode(true));
                }
            }

            AddSectionProperties(sectionProperties);
            return sectionProperties;
        }

        internal SectionProperties CreateDefaultSectionProperties(bool addToDocument = true)
        {
            var sectionProperties = new SectionProperties();

            // Set default RTF 1.9.1 properties unless already set
            SetMissingSectionProperty<PageSize>(sectionProperties, pgSize =>
            {
                pgSize.Width = 12240;  
                pgSize.Height = 15840; 
            });
            SetMissingSectionProperty<PageMargin>(sectionProperties, pgMargin =>
            {
                pgMargin.Left = 1800;   
                pgMargin.Right = 1800;  
                pgMargin.Top = 1440;    
                pgMargin.Bottom = 1440;
                pgMargin.Header = 720;
                pgMargin.Footer = 720;
            });

            if (addToDocument)
            {
                AddDefaultSectionProperties(sectionProperties);
            }
            return sectionProperties;
        }

        internal void AddSectionProperties(SectionProperties sectionProperties)
        {
            if (sectionProperties.Parent == null)
            {
                // Insert the new section properties before the last section properties (which is the default)
                if (_body.Elements<SectionProperties>().LastOrDefault() is SectionProperties lastSection)
                {
                    if (lastSection != sectionProperties)
                    {
                        _body.InsertBefore(sectionProperties, lastSection);
                    }
                }
                else
                {
                    // If the last section properties is not present, append the new one
                    _body.AppendChild(sectionProperties);
                }
            }
        }

        internal void AddDefaultSectionProperties(SectionProperties sectionProperties, bool forceReplace = true)
        {
            if (sectionProperties.Parent != null)
                return;

            if (_body.Elements().LastOrDefault() is SectionProperties existing)
            {
                if (forceReplace)
                {
                    existing.Remove();
                    _body.AppendChild(sectionProperties);
                }
            }
            else
            {
                _body.AppendChild(sectionProperties);
            }
        }

        internal void SetSectionProperty<T>(SectionProperties sectionProperties, Action<T> setProperty) where T : OpenXmlElement, new()
        {
            var element = sectionProperties.Elements<T>().FirstOrDefault();
            if (element == null)
            {
                element = new T();
                sectionProperties.AppendChild(element);
            }
            setProperty(element);
        }

        internal void SetPageDimensions(SectionProperties sectionProperties, int widthOrHeight, bool isWidth)
        {
            SetSectionProperty<PageSize>(sectionProperties, pgSize =>
            {
                if (isWidth)
                    pgSize.Width = (uint)widthOrHeight;
                else
                    pgSize.Height = (uint)widthOrHeight;
            });
        }

        internal void SetPageMargin(SectionProperties sectionProperties, int value, string marginType)
        {
            SetSectionProperty<PageMargin>(sectionProperties, pgMargin =>
            {
                switch (marginType)
                {
                    case "Left":
                        pgMargin.Left = (uint)value;
                        break;
                    case "Right":
                        pgMargin.Right = (uint)value;
                        break;
                    case "Top":
                        pgMargin.Top = value;
                        break;
                    case "Bottom":
                        pgMargin.Bottom = value;
                        break;
                    case "Gutter":
                        pgMargin.Gutter = (uint)value;
                        break;
                    case "Header":
                        pgMargin.Header = (uint)value;
                        break;
                    case "Footer":
                        pgMargin.Footer = (uint)value;
                        break;
                }
            });
        }

        internal void SetSectionType(SectionProperties sectionProperties, SectionMarkValues sectionType)
        {
            SetSectionProperty<SectionType>(sectionProperties, sectionTypeElement =>
            {
                sectionTypeElement.Val = sectionType;
            });
        }

        internal ParagraphProperties CreateDefaultParagraphProperties()
        {
            var paragraphProperties = new ParagraphProperties();

            // Set default properties unless already set
            //SetMissingParagraphProperty<Justification>(paragraphProperties, justification =>
            //{
            //    justification.Val = JustificationValues.Left; // Default alignment
            //});

            return paragraphProperties;
        }

        internal void SetParagraphProperty<T>(Action<T> setProperty) where T : OpenXmlElement, new()
        {
            EnsureCurrentParagraph();
            if (_currentParagraphProperties == null)
            {
                _currentParagraphProperties = new ParagraphProperties();
                _currentParagraph!.PrependChild(_currentParagraphProperties);
            }
            SetParagraphProperty(_currentParagraphProperties, setProperty);
        }

        internal void SetParagraphProperty<T>(ParagraphProperties paragraphProperties, Action<T> setProperty) where T : OpenXmlElement, new()
        {
            var element = paragraphProperties.Elements<T>().FirstOrDefault();
            if (element == null)
            {
                element = new T();
                paragraphProperties.AppendChild(element);
            }
            setProperty(element);
        }

        internal void SetParagraphAlignment(ParagraphProperties paragraphProperties, JustificationValues alignment)
        {
            SetParagraphProperty<Justification>(paragraphProperties, justification =>
            {
                justification.Val = alignment;
            });
        }

        internal void SetParagraphIndentation(ParagraphProperties paragraphProperties, int left, int right, int firstLine)
        {
            SetParagraphProperty<Indentation>(paragraphProperties, indentation =>
            {
                indentation.Left = left.ToString();
                indentation.Right = right.ToString();
                indentation.FirstLine = firstLine.ToString();
            });
        }

        internal void SetMissingSectionProperty<T>(SectionProperties sectionProperties, Action<T> setProperty) where T : OpenXmlElement, new()
        {
            var element = sectionProperties.Elements<T>().FirstOrDefault();
            if (element == null)
            {
                element = new T();
                sectionProperties.AppendChild(element);
                setProperty(element);
            }
        }

        internal void SetMissingParagraphProperty<T>(ParagraphProperties paragraphProperties, Action<T> setProperty) where T : OpenXmlElement, new()
        {
            var element = paragraphProperties.Elements<T>().FirstOrDefault();
            if (element == null)
            {
                element = new T();
                paragraphProperties.AppendChild(element);
                setProperty(element);
            }
        }

        internal void AddContentElement(OpenXmlElement? element)
        {
            if (element != null)
            {
                if (_currentSectionProperties != null)
                {
                    if (element is Paragraph paragraph || element is Table)
                    {
                        AddSectionProperties(_currentSectionProperties);
                        _body.InsertBefore(element, _currentSectionProperties);
                    }
                }
                else if (_defaultSectionProperties != null)
                {
                    if (element is Paragraph paragraph || element is Table)
                    {
                        AddDefaultSectionProperties(_defaultSectionProperties, false);
                        _body.InsertBefore(element, _defaultSectionProperties);
                    }
                }
                else if (element is SectionProperties)
                {
                    // Should not happen
                    _body.AppendChild(element);
                }
                else
                {
                    // Should not happen
                    _defaultSectionProperties = CreateDefaultSectionProperties();
                    if (element is Paragraph paragraph || element is Table)
                    {
                        _body.InsertBefore(element, _defaultSectionProperties);
                    }
                }
            }
        }

        private void EnsureCurrentRun()
        {
            EnsureCurrentParagraph();
            if (_currentRun == null)
            {
                _currentRun = new Run();
                _currentParagraph!.AppendChild(_currentRun);
            }
        }

        private void EnsureCurrentParagraph()
        {
            if (_currentParagraph == null)
            {
                _currentParagraph = new Paragraph();

                if (_currentParagraphProperties != null)
                {
                    _currentParagraph.AppendChild(_currentParagraphProperties.CloneNode(true));
                }
                else
                {
                    _currentParagraph.AppendChild(CreateDefaultParagraphProperties());
                }

                AddContentElement(_currentParagraph);
            }
        }
    }
}
