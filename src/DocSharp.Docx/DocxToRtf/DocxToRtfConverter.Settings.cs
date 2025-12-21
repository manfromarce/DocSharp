using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal void ProcessSettings(Settings? settings, RtfStringWriter writer)
    {
        if (settings?.GetFirstChild<DocumentType>() is DocumentType docType && docType.Val != null)
        {
            if (docType.Val.Value == DocumentTypeValues.NotSpecified)
                writer.Write(@"\doctype0");
            else if (docType.Val.Value == DocumentTypeValues.Letter)
                writer.Write(@"\doctype1");
            else if (docType.Val.Value == DocumentTypeValues.Email)
                writer.Write(@"\doctype2");
        }

        // These control words have the same meaning and default value in DOCX and RTF, so we can avoid writing them if false or not present.
        // They are written followed by "1" if true.
        writer.WriteWordWithValueIfTrue("hyphauto", (settings?.GetFirstChild<AutoHyphenation>()).ToBool());
        writer.WriteWordWithValueIfTrue("saveinvalidxml", (settings?.GetFirstChild<SaveInvalidXml>()).ToBool());
        writer.WriteWordWithValueIfTrue("ignoremixedcontent", (settings?.GetFirstChild<IgnoreMixedContent>()).ToBool());
        writer.WriteWordWithValueIfTrue("showplaceholdtext", (settings?.GetFirstChild<AlwaysShowPlaceholderText>()).ToBool());
        writer.WriteWordWithValueIfTrue("xform", (settings?.GetFirstChild<UseXsltWhenSaving>()).ToBool());
        //if (settings?.GetFirstChild<SaveThroughXslt>() is SaveThroughXslt saveThroughXslt && saveThroughXslt.SolutionId?.Value != null)
        //{
        // TODO: retrieve the file path
        //writer.Write(@$"{{\*\xform {path}}}"); // e.g.: {\*\xform c:\\temp\\myxslt.xsl}
        //}

        // These control words should **not** be followed by "1"; they mean "true" if present, otherwise we don't write them.
        writer.WriteWordIfTrue("pgbrdrsnap", (settings?.AlignBorderAndEdges).ToBool());
        writer.WriteWordIfTrue("autofmtoverride", (settings?.GetFirstChild<AutoFormatOverride>()).ToBool());
        writer.WriteWordIfTrue("bookfold", (settings?.GetFirstChild<BookFoldPrinting>()).ToBool());
        writer.WriteWordIfTrue("bookfoldrev", (settings?.GetFirstChild<BookFoldReversePrinting>()).ToBool());
        writer.WriteWordIfTrue("gutterprl", (settings?.GutterAtTop).ToBool());
        writer.WriteWordIfTrue("linkstyles", (settings?.GetFirstChild<LinkStyles>()).ToBool());
        writer.WriteWordIfTrue("margmirror", (settings?.MirrorMargins).ToBool());
        writer.WriteWordIfTrue("nojkernpunct", (settings?.GetFirstChild<NoPunctuationKerning>()).ToBool());
        writer.WriteWordIfTrue("fracwidth", (settings?.PrintFractionalCharacterWidth).ToBool());
        writer.WriteWordIfTrue("printdata", (settings?.PrintFormsData).ToBool());
        writer.WriteWordIfTrue("psover", (settings?.PrintPostScriptOverText).ToBool());
        writer.WriteWordIfTrue("twoonone", (settings?.GetFirstChild<PrintTwoOnOne>()).ToBool());
        writer.WriteWordIfTrue("remdttm", (settings?.RemoveDateAndTime).ToBool());
        writer.WriteWordIfTrue("rempersonalinfo", (settings?.RemovePersonalInformation).ToBool());
        writer.WriteWordIfTrue("saveprevpict", (settings?.GetFirstChild<SavePreviewPicture>()).ToBool());
        writer.WriteWordIfTrue("stylelockqfset", (settings?.GetFirstChild<StyleLockStylesPart>()).ToBool());
        writer.WriteWordIfTrue("stylelocktheme", (settings?.GetFirstChild<StyleLockThemesPart>()).ToBool());
        writer.WriteWordIfTrue("revisions", (settings?.GetFirstChild<TrackRevisions>()).ToBool());
        writer.WriteWordIfTrue("jsksu", (settings?.GetFirstChild<StrictFirstAndLastChars>()).ToBool());

        // These control words have the opposite meaning in RTF;
        // the default value in DOCX is still false (which means show xml errors, **don't** embed system fonts...),
        // in RTF it it may vary so always write the value for clarity
        writer.WriteWordWithValue("showxmlerrors", !(settings?.GetFirstChild<DoNotDemarcateInvalidXml>()).ToBool());
        writer.WriteWordWithValue("validatexml", !(settings?.GetFirstChild<DoNotValidateAgainstSchema>()).ToBool());
        writer.WriteWordWithValue("donotembedsysfont", !(settings?.GetFirstChild<EmbedSystemFonts>()).ToBool());
        writer.WriteWordWithValue("hyphcaps", !(settings?.GetFirstChild<DoNotHyphenateCaps>()).ToBool());
        writer.WriteWordWithValue("trackformatting", !(settings?.GetFirstChild<DoNotTrackFormatting>()).ToBool());
        writer.WriteWordWithValue("trackmoves", !(settings?.GetFirstChild<DoNotTrackMoves>()).ToBool());
        
        // These control words have the opposite meaning in RTF, are false by default in DOCX, and should be written without "1"
        writer.WriteWordIfTrue("formshade", !((settings?.GetFirstChild<DoNotShadeFormData>()).ToBool()));
        writer.WriteWordIfTrue("dgmargin", !(settings?.GetFirstChild<DoNotUseMarginsForDrawingGridOrigin>()).ToBool());
        writer.WriteWordIfTrue("pgbrdrhead", !((settings?.BordersDoNotSurroundHeader).ToBool()));
        writer.WriteWordIfTrue("pgbrdrfoot", !((settings?.BordersDoNotSurroundFooter).ToBool()));

        // These are numeric values
        if (settings?.GetFirstChild<DefaultTabStop>() is DefaultTabStop defaultTabStop && defaultTabStop.Val != null)
            writer.WriteWordWithValue("deftab", defaultTabStop.Val.Value);
        if (settings?.GetFirstChild<BookFoldPrintingSheets>() is BookFoldPrintingSheets sheets && sheets.Val != null)
            writer.WriteWordWithValue("bookfoldsheets", sheets.Val.Value);
        if (settings?.GetFirstChild<HyphenationZone>() is HyphenationZone hz && hz.Val.ToLong() is long hyphZone)
            writer.WriteWordWithValue("hyphhotz", hyphZone);
        if (settings?.GetFirstChild<ConsecutiveHyphenLimit>() is ConsecutiveHyphenLimit consecutiveHyphenLimit && consecutiveHyphenLimit.Val != null)
            writer.WriteWordWithValue("hyphconsec", consecutiveHyphenLimit.Val);
        if (settings?.GetFirstChild<DrawingGridHorizontalSpacing>() is DrawingGridHorizontalSpacing dghs && dghs.Val.ToLong() is long dghsVal)
            writer.WriteWordWithValue("dghspace", dghsVal);
        if (settings?.GetFirstChild<DrawingGridVerticalSpacing>() is DrawingGridVerticalSpacing dgvs && dgvs.Val.ToLong() is long dgvsVal)
            writer.WriteWordWithValue("dgvspace", dgvsVal);
        if (settings?.GetFirstChild<DrawingGridHorizontalOrigin>() is DrawingGridHorizontalOrigin dgho && dgho.Val.ToLong() is long dghoVal)
            writer.WriteWordWithValue("dghorigin", dghoVal);
        if (settings?.GetFirstChild<DrawingGridVerticalOrigin>() is DrawingGridVerticalOrigin dgvo && dgvo.Val.ToLong() is long dgvoVal)
            writer.WriteWordWithValue("dgvorigin", dgvoVal);
        if (settings?.GetFirstChild<DisplayHorizontalDrawingGrid>() is DisplayHorizontalDrawingGrid dhdg && dhdg.Val != null)
            writer.WriteWordWithValue("dghshow", dhdg.Val.Value);
        if (settings?.GetFirstChild<DisplayVerticalDrawingGrid>() is DisplayVerticalDrawingGrid dvdg && dvdg.Val != null)
            writer.WriteWordWithValue("dgvshow", dvdg.Val.Value);

        if (settings?.WriteProtection is WriteProtection wp)
        {
            if (wp.Recommended != null && wp.Recommended)
                writer.Write(@"\readonlyrecommended");
            // TODO: other attributes (\*\passwordhash and \*\writereservhash destinations;
            // \formprot, \revprot, \annotprot, \readprot)
        }

        if (settings?.View is View view)
        {
            if (view.Val != null)
            {
                if (view.Val == ViewValues.None)
                    writer.WriteWordWithValue("viewkind", 0);
                else if (view.Val == ViewValues.Print)
                    writer.WriteWordWithValue("viewkind", 1);
                else if (view.Val == ViewValues.Outline)
                    writer.WriteWordWithValue("viewkind", 2);
                else if (view.Val == ViewValues.MasterPages)
                    writer.WriteWordWithValue("viewkind", 3);
                else if (view.Val == ViewValues.Normal)
                    writer.WriteWordWithValue("viewkind", 4);
                else if (view.Val == ViewValues.Web)
                    writer.WriteWordWithValue("viewkind", 5);
            }
        }

        if (settings?.Zoom is Zoom zoom)
        {
            if (zoom.Val != null)
            {
                if (zoom.Val == PresetZoomValues.FullPage)
                    writer.WriteWordWithValue("viewzk", 1);
                else if (zoom.Val == PresetZoomValues.BestFit)
                    writer.WriteWordWithValue("viewzk", 2);
                else if (zoom.Val == PresetZoomValues.TextFit)
                    writer.WriteWordWithValue("viewzk", 3);
            }
            if ((zoom.Val == null || zoom.Val == PresetZoomValues.None) &&
                zoom.Percent?.Value != null &&
                decimal.TryParse(zoom.Percent.Value.TrimEnd('%'), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal zoomPercent))
            {
                writer.WriteWordWithValue("viewzk", 0);
                writer.WriteWordWithValue("viewscale", (long)Math.Round(zoomPercent));
            }
        }

        // TODO
        //if (settings?.GetFirstChild<DocumentProtection>() is DocumentProtection dp)
        //{
        //
        //}

        if (settings?.GetFirstChild<CharacterSpacingControl>() is CharacterSpacingControl csp && csp.Val != null && 
            csp.Val != CharacterSpacingValues.DoNotCompress)
            writer.Write("\\jcompress");
        else
            writer.Write("\\jexpand"); // default

        ProcessCompatibilityOptions(settings?.GetFirstChild<Compatibility>(), writer);

        ProcessMathDocumentProperties(settings?.GetFirstChild<M.MathProperties>(), writer);
        ProcessFootnoteProperties(settings?.GetFirstChild<FootnoteDocumentWideProperties>(), writer);
        ProcessEndnoteProperties(settings?.GetFirstChild<EndnoteDocumentWideProperties>(), writer);
        ProcessFacingPages(settings?.GetFirstChild<EvenAndOddHeaders>(), writer);
    }

    internal void ProcessCompatibilityOptions(Compatibility? compat, RtfStringWriter writer)
    {
        // Note: the default values for these are often counterintuitive (e.g. CachedColumnBalance, UseWord2002TableStyleRules)
        // and were determined by enabling compatibility options in Word and comparing documents saved as DOCX and RTF.
        if (compat?.UseSingleBorderForContiguousCells != null && (compat.UseSingleBorderForContiguousCells.Val == null || compat.UseSingleBorderForContiguousCells.Val))
        {
            writer.Write(@"\otblrul"); //?
        }
        if (compat?.WordPerfectJustification != null && (compat.WordPerfectJustification.Val == null || compat.WordPerfectJustification.Val))
        {
            writer.Write(@"\wpjst");
        }
        if (compat?.NoTabHangIndent != null && (compat.NoTabHangIndent.Val == null || compat.NoTabHangIndent.Val))
        {
            writer.Write(@"\notabind");
        }
        if (compat?.NoLeading != null && (compat.NoLeading.Val == null || compat.NoLeading.Val))
        {
            writer.Write(@"\nolead");
        }
        if (compat?.NoColumnBalance != null && (compat.NoColumnBalance.Val == null || compat.NoColumnBalance.Val))
        {
            writer.Write(@"\nocolbal");
        }
        if (compat?.LineWrapLikeWord6 != null && (compat.LineWrapLikeWord6.Val == null || compat.LineWrapLikeWord6.Val))
        {
            writer.Write(@"\oldlinewrap");
        }
        if (compat?.PrintBodyTextBeforeHeader != null && (compat.PrintBodyTextBeforeHeader.Val == null || compat.PrintBodyTextBeforeHeader.Val))
        {
            writer.Write(@"\bdbfhdr");
        }
        if (compat?.PrintColorBlackWhite != null && (compat.PrintColorBlackWhite.Val == null || compat.PrintColorBlackWhite.Val))
        {
            writer.Write(@"\prcolbl");
        }
        if (compat?.ShowBreaksInFrames != null && (compat.ShowBreaksInFrames.Val == null || compat.ShowBreaksInFrames.Val))
        {
            writer.Write(@"\brkfrm");
        }
        if (compat?.SubFontBySize != null && (compat.SubFontBySize.Val == null || compat.SubFontBySize.Val))
        {
            writer.Write(@"\subfontbysize");
        }
        if (compat?.SuppressTopSpacingWordPerfect != null && (compat.SuppressTopSpacingWordPerfect.Val == null || compat.SuppressTopSpacingWordPerfect.Val))
        {
            writer.Write(@"\sprslnsp");
        }
        if (compat?.SuppressSpacingBeforeAfterPageBreak != null && (compat.SuppressSpacingBeforeAfterPageBreak.Val == null || compat.SuppressSpacingBeforeAfterPageBreak.Val))
        {
            writer.Write(@"\sprsspbf");
        }
        if (compat?.SwapBordersFacingPages != null && (compat.SwapBordersFacingPages.Val == null || compat.SwapBordersFacingPages.Val))
        {
            writer.Write(@"\swpbdr");
        }
        if (compat?.ConvertMailMergeEscape != null && (compat.ConvertMailMergeEscape.Val == null || compat.ConvertMailMergeEscape.Val))
        {
            writer.Write(@"\cvmme");
        }
        if (compat?.TruncateFontHeightsLikeWordPerfect != null && (compat.TruncateFontHeightsLikeWordPerfect.Val == null || compat.TruncateFontHeightsLikeWordPerfect.Val))
        {
            writer.Write(@"\truncatefontheight");
        }
        if (compat?.MacWordSmallCaps != null && (compat.MacWordSmallCaps.Val == null || compat.MacWordSmallCaps.Val))
        {
            writer.Write(@"\msmcap");
        }
        if (compat?.UsePrinterMetrics != null && (compat.UsePrinterMetrics.Val == null || compat.UsePrinterMetrics.Val))
        {
            writer.Write(@"\lytprtmet");
        }
        if (compat?.DoNotSuppressParagraphBorders != null && (compat.DoNotSuppressParagraphBorders.Val == null || compat.DoNotSuppressParagraphBorders.Val))
        {
            writer.Write(@"\bdrrlswsix"); // ?
        }
        if (compat?.WrapTrailSpaces != null && (compat.WrapTrailSpaces.Val == null || compat.WrapTrailSpaces.Val))
        {
            writer.Write(@"\wraptrsp");
        }
        if (compat?.AutoSpaceLikeWord95 != null && (compat.AutoSpaceLikeWord95.Val == null || compat.AutoSpaceLikeWord95.Val))
        {
            writer.Write(@"\oldas");
        }
        if (compat?.WordPerfectSpaceWidth != null && (compat.WordPerfectSpaceWidth.Val == null || compat.WordPerfectSpaceWidth.Val))
        {
            writer.Write(@"\wpsp");
        }
        if (compat?.SuppressBottomSpacing != null && (compat.SuppressBottomSpacing.Val == null || compat.SuppressBottomSpacing.Val))
        {
            writer.Write(@"\sprsbsp");
        }
        if (compat?.SuppressTopSpacing != null && (compat.SuppressTopSpacing.Val == null || compat.SuppressTopSpacing.Val))
        {
            writer.Write(@"\sprstsp");
        }
        if (compat?.SuppressSpacingAtTopOfPage != null && (compat.SuppressSpacingAtTopOfPage.Val == null || compat.SuppressSpacingAtTopOfPage.Val))
        {
            writer.Write(@"\sprstsm");
        }
        if (compat?.NoSpaceRaiseLower != null && (compat.NoSpaceRaiseLower.Val == null || compat.NoSpaceRaiseLower.Val))
        {
            writer.Write(@"\noextrasprl");
        }
        if (compat?.ApplyBreakingRules != null && (compat.ApplyBreakingRules.Val == null || compat.ApplyBreakingRules.Val))
        {
            writer.Write(@"\ApplyBrkRules"); // ?
        }
        if (compat?.NoExtraLineSpacing != null && (compat.NoExtraLineSpacing.Val == null || compat.NoExtraLineSpacing.Val))
        {
            writer.Write(@"\lytexcttp");
        }
        if (compat?.SpacingInWholePoints != null && (compat.SpacingInWholePoints.Val == null || compat.SpacingInWholePoints.Val))
        {
            writer.Write(@"\truncex");
        }
        //if (compat?.UseFarEastLayout != null && (compat.UseFarEastLayout.Val == null || compat.UseFarEastLayout.Val))
        //{
        //}

        // ----

        if (compat?.DoNotLeaveBackslashAlone == null || (compat.DoNotLeaveBackslashAlone.Val != null && compat.DoNotLeaveBackslashAlone.Val == false))
        {
            writer.Write(@"\noxlattoyen");
        }
        if (compat?.BalanceSingleByteDoubleByteWidth == null || (compat.BalanceSingleByteDoubleByteWidth.Val != null && compat.BalanceSingleByteDoubleByteWidth.Val == false))
        {
            writer.Write(@"\dntblnsbdb");
        }
        if (compat?.AdjustLineHeightInTable == null || (compat.AdjustLineHeightInTable.Val != null && compat.AdjustLineHeightInTable.Val == false))
        {
            writer.Write(@"\nolnhtadjtbl");
        }
        if (compat?.DoNotExpandShiftReturn == null || (compat.DoNotExpandShiftReturn.Val != null && compat.DoNotExpandShiftReturn.Val == false))
        {
            writer.Write(@"\expshrtn");
        }
        if (compat?.UnderlineTrailingSpaces == null || (compat.UnderlineTrailingSpaces.Val != null && compat.UnderlineTrailingSpaces.Val == false))
        {
            writer.Write(@"\noultrlspc");
        }
        if (compat?.SpaceForUnderline == null || (compat.SpaceForUnderline.Val != null && compat.SpaceForUnderline.Val == false))
        {
            writer.Write(@"\nospaceforul");
        }
        if (compat?.UnderlineTabInNumberingList == null || (compat.UnderlineTabInNumberingList.Val != null && compat.UnderlineTabInNumberingList.Val == false))
        {
            writer.Write(@"\utinl");
        }
        if (compat?.DoNotUseHTMLParagraphAutoSpacing == null || (compat.DoNotUseHTMLParagraphAutoSpacing.Val != null && compat.DoNotUseHTMLParagraphAutoSpacing.Val == false))
        {
            writer.Write(@"\htmautsp");
        }
        if (compat?.ForgetLastTabAlignment == null || (compat.ForgetLastTabAlignment.Val != null && compat.ForgetLastTabAlignment.Val == false))
        {
            writer.Write(@"\useltbaln");
        }
        if (compat?.UseWord97LineBreakRules == null || (compat.UseWord97LineBreakRules.Val != null && compat.UseWord97LineBreakRules.Val == false))
        {
            writer.Write(@"\lnbrkrule");
        }
        if (compat?.CachedColumnBalance == null || (compat.CachedColumnBalance.Val != null && compat.CachedColumnBalance.Val == false))
        {
            writer.Write(@"\cachedcolbal");
        }
        if (compat?.DoNotAutofitConstrainedTables == null || (compat.DoNotAutofitConstrainedTables.Val != null && compat.DoNotAutofitConstrainedTables.Val == false))
        {
            writer.Write(@"\noafcnsttbl");
        }
        if (compat?.DisplayHangulFixedWidth == null || (compat.DisplayHangulFixedWidth.Val != null && compat.DisplayHangulFixedWidth.Val == false))
        {
            writer.Write(@"\hwelev");
        }
        if (compat?.SplitPageBreakAndParagraphMark == null || (compat.SplitPageBreakAndParagraphMark.Val != null && compat.SplitPageBreakAndParagraphMark.Val == false))
        {
            writer.Write(@"\spltpgpar");
        }
        if (compat?.DoNotVerticallyAlignCellWithShape == null || (compat.DoNotVerticallyAlignCellWithShape.Val != null && compat.DoNotVerticallyAlignCellWithShape.Val == false))
        {
            writer.Write(@"\notcvasp");
        }
        if (compat?.DoNotVerticallyAlignInTextBox == null || (compat.DoNotVerticallyAlignInTextBox.Val != null && compat.DoNotVerticallyAlignInTextBox.Val == false))
        {
            writer.Write(@"\notvatxbx");
        }
        if (compat?.DoNotBreakConstrainedForcedTable == null || (compat.DoNotBreakConstrainedForcedTable.Val != null && compat.DoNotBreakConstrainedForcedTable.Val == false))
        {
            writer.Write(@"\notbrkcnstfrctbl");
        }
        if (compat?.DoNotBreakWrappedTables == null || (compat.DoNotBreakWrappedTables.Val != null && compat.DoNotBreakWrappedTables.Val == false))
        {
            writer.Write(@"\nobrkwrptbl");
        }
        if (compat?.UseAnsiKerningPairs == null || (compat.UseAnsiKerningPairs.Val != null && compat.UseAnsiKerningPairs.Val == false))
        {
            writer.Write(@"\krnprsnet");
        }
        if (compat?.UseAltKinsokuLineBreakRules == null || (compat.UseAltKinsokuLineBreakRules.Val != null && compat.UseAltKinsokuLineBreakRules.Val == false))
        {
            writer.Write(@"\felnbrelev");
        }
        if (compat?.DoNotSuppressIndentation == null || (compat.DoNotSuppressIndentation.Val != null && compat.DoNotSuppressIndentation.Val == false))
        {
            writer.Write(@"\indrlsweleven"); // ?
        }
        if (compat?.DoNotSnapToGridInCell == null || (compat.DoNotSnapToGridInCell.Val != null && compat.DoNotSnapToGridInCell.Val == false))
        {
            writer.Write(@"\snaptogridincell");
        }
        if (compat?.SelectFieldWithFirstOrLastChar == null || (compat.SelectFieldWithFirstOrLastChar.Val != null && compat.SelectFieldWithFirstOrLastChar.Val == false))
        {
            writer.Write(@"\allowfieldendsel");
        }
        if (compat?.DoNotWrapTextWithPunctuation == null || (compat.DoNotWrapTextWithPunctuation.Val != null && compat.DoNotWrapTextWithPunctuation.Val == false))
        {
            writer.Write(@"\wrppunct");
        }
        if (compat?.DoNotUseEastAsianBreakRules == null || (compat.DoNotUseEastAsianBreakRules.Val != null && compat.DoNotUseEastAsianBreakRules.Val == false))
        {
            writer.Write(@"\asianbrkrule");
        }
        if (compat?.UseWord2002TableStyleRules == null || (compat.UseWord2002TableStyleRules.Val != null && compat.UseWord2002TableStyleRules.Val == false))
        {
            writer.Write(@"\newtblstyruls");
        }
        if (compat?.GrowAutofit == null || (compat.GrowAutofit.Val != null && compat.GrowAutofit.Val == false))
        {
            writer.Write(@"\nogrowautofit");
        }
        if (compat?.DoNotUseIndentAsNumberingTabStop == null || (compat.DoNotUseIndentAsNumberingTabStop.Val != null && compat.DoNotUseIndentAsNumberingTabStop.Val == false))
        {
            writer.Write(@"\noindnmbrts");
        }
        if (compat?.UseNormalStyleForList == null || (compat.UseNormalStyleForList.Val != null && compat.UseNormalStyleForList.Val == false))
        {
            writer.Write(@"\usenormstyforlist");
        }
        if (compat?.AllowSpaceOfSameStyleInTable == null || (compat.AllowSpaceOfSameStyleInTable.Val != null && compat.AllowSpaceOfSameStyleInTable.Val == false))
        {
            writer.Write(@"\nocxsptable");
        }
        if (compat?.LayoutRawTableWidth == null || (compat.LayoutRawTableWidth.Val != null && compat.LayoutRawTableWidth.Val == false))
        {
            writer.Write(@"\lytcalctblwd");
        }
        if (compat?.LayoutTableRowsApart == null || (compat.LayoutTableRowsApart.Val != null && compat.LayoutTableRowsApart.Val == false))
        {
            writer.Write(@"\lyttblrtgr");
        }
        if (compat?.AlignTablesRowByRow == null || (compat.AlignTablesRowByRow.Val != null && compat.AlignTablesRowByRow.Val == false))
        {
            writer.Write(@"\alntblind");
        }
        if (compat?.FootnoteLayoutLikeWord8 == null || (compat.FootnoteLayoutLikeWord8.Val != null && compat.FootnoteLayoutLikeWord8.Val == false))
        {
            writer.Write(@"\ftnlytwnine");
        }
        if (compat?.ShapeLayoutLikeWord8 == null || (compat.ShapeLayoutLikeWord8.Val != null && compat.ShapeLayoutLikeWord8.Val == false))
        {
            writer.Write(@"\splytwnine");
        }
        if (compat?.AutofitToFirstFixedWidthCell == null || (compat.AutofitToFirstFixedWidthCell.Val != null && compat.AutofitToFirstFixedWidthCell.Val == false))
        {
            writer.Write(@"\afelev"); // ?
        }

        //if (compat?.Elements<CompatibilitySetting>().Where(cs => cs.Name != null && cs.Name == CompatSettingNameValues.UseWord2013TrackBottomHyphenation).FirstOrDefault() is CompatibilitySetting setting)
        //{
        //}
    }
}
