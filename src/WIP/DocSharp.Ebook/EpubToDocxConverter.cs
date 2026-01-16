using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocSharp.Helpers;
using DocSharp.Primitives;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using EpubCore;
using HtmlToOpenXml;
using DocSharp.Docx;

namespace DocSharp.Ebook;

// NOTE: internal links currently might throw exceptions in the Html2OpenXml library in some cases; 
// we should also detect the cover and make it a full-page image.

/// <summary>
/// Basic experimental EPUB to DOCX converter that performs the following steps: 
/// 1. Extract the EPUB into a temp folder
/// 2. Reads the EPUB using EpubCore to get the chapters in reading order
/// 3. For each chapter, replace image sources with absolute URIs, 
/// attempt to fix links to other chapters, 
/// and move CSS styles inline using the PreMailer library.
/// 4. Convert HTML to DOCX using the HtmlToOpenXml library and append to the DOCX document.
/// 5. Delete the temp directory.
/// </summary>
public class EpubToDocxConverter : IBinaryToDocxConverter
{
    /// <summary>
    /// If true, only the "core" chapters will get converted.
    /// The default is false, thus including cover, table of contents and other transition pages in the output document.
    /// </summary>
    public bool ChaptersOnly { get; set; } = false;

    /// <summary>
    /// If true, adds a page break after each chapter.
    /// </summary>
    public bool PageBreakAfterChapters { get; set; } = true;

    /// <summary>
    /// If true, the converter will attempt to preserve CSS styles embedded in the eBook and create equivalents in the output DOCX document.
    /// The default is true, set to false to produce a more minimal document if unexpected/undesired formatting is present.
    /// </summary>
    public bool PreserveCssStyles { get; set; } = true;

    /// <summary>
    /// Page size for the output DOCX document.
    /// </summary>
    public PageSize PageSize { get; set; } = PageSize.Default;

    /// <summary>
    /// Page margins for the output DOCX document.
    /// </summary>
    public PageMargins PageMargins { get; set; } = PageMargins.Default;

    public async Task BuildDocxAsync(Stream input, WordprocessingDocument targetDocument)
    {
        // Create temp directory
        var tempDir = Path.Combine(Path.GetTempPath(), "epub_extract_" + Path.GetRandomFileName());
        if (!tempDir.EndsWith(Path.DirectorySeparatorChar))
            // Add the final slash, as the PreMailer library has issues in finding resources.
            tempDir += Path.DirectorySeparatorChar; 
        try
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir);
            Directory.CreateDirectory(tempDir);            
        }
        catch (Exception ex)
        {
            throw new SystemException($"Unable to create a temp directory. This step is necessary for EPUB processing. Details: {ex.Message}");
        }

        try
        {
            // Extract EPUB to temp directory
            using (var zip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
            {
                zip.ExtractToDirectory(tempDir);
            }
        }
        catch (Exception ex)
        {
            throw new SystemException($"Unable to extract EPUB to the temp directory. This step is necessary for EPUB processing. Details: {ex.Message}");
        }

        try
        {
            // Read EPUB
            var book = EpubReader.Read(input, leaveOpen: true);

            // Get chapters (or all html pages including cover and table of contents), 
            // depending on the ChaptersOnly property.
            var chapters = ChaptersOnly ? book.TableOfContents.Select(chapter => book.FetchHtmlFileForChapter(chapter)).ToList() : 
                                        book.SpecialResources.HtmlInReadingOrder.ToList();
            var chapterFileNames = chapters.Select(file => file.FileName).ToList();

            // Initialize document
            var mainPart = targetDocument.MainDocumentPart ?? targetDocument.AddMainDocumentPart();
            mainPart.Document ??= new W.Document();
            mainPart.Document.RemoveAllChildren();
            var body = mainPart.Document.AppendChild(new W.Body());

            // Initialize HTML to DOCX converter
            var converter = new HtmlConverter(mainPart)
            {
                ImageProcessing = ImageProcessingMode.Embed,
                SupportsAnchorLinks = true,
                SupportsHeadingNumbering = true
            };

            // Enumerate chapters
            foreach (var chapter in chapters)
            {
                try
                {
                    // Get chapter file name and XHTML content
                    var fileName = chapter.FileName;
                    var htmlContent = chapter.TextContent;

                    // Attempt to fix external images sources and links pointing to other chapters. 
                    var normalizedHtml = HtmlUtils.NormalizeHtml(htmlContent, tempDir, chapterFileNames);

                    // HtmlToOpenXml can load external images, while styles should be moved inline: 
                    // https://github.com/onizet/html2openxml/wiki/Style
                    // Move styles inline using the PreMailer.Net library, unless style conversion is disabled.
                    var htmlWithInlinedCss = normalizedHtml;
                    if (PreserveCssStyles)
                    {   
                        try
                        {                        
                            var inliner = new PreMailer.Net.PreMailer(normalizedHtml, new Uri(tempDir));
                            var inlinerResult = inliner.MoveCssInline(
                                removeStyleElements: true, 
                                stripIdAndClassAttributes: true
                            );
                            htmlWithInlinedCss = inlinerResult.Html;                    
                        }
                        catch(Exception ex)
                        {
                            #if DEBUG
                            Debug.WriteLine($"Error during style inlining. Details: {ex.Message}");
                            #endif
                        }
                    }

                    // Before each chapter, add a bookmark in DOCX to make internal links work
                    string anchorName = $"_{fileName.Replace(' ', '_')}";
                    int id = new Random().Next(100000, 999999); // TODO: improve id generation
                    body.AppendChild(new W.Paragraph([
                        new W.BookmarkStart() { Name = anchorName, Id = id.ToString() }, 
                        new W.BookmarkEnd() { Id = id.ToString() }
                    ]));

                    // Parse the HTML body, convert to Open XML and append to the DOCX.
                    await converter.ParseBody(htmlWithInlinedCss);                        

                    if (PageBreakAfterChapters)
                    {
                        // Add a page break after each chapter if desired.
                        body.AppendChild(new W.Paragraph(new W.Run(new W.Break() { Type = W.BreakValues.Page })));
                    }
                }
                catch (Exception chapterException)
                {
                    #if DEBUG
                    Debug.WriteLine($"Exception in chapter: {chapter.FileName} - {chapterException.Message}");
                    continue;
                    #endif
                }
            }            

            // Add default section properties            
            body.AppendChild(new W.SectionProperties(
                new W.PageSize()
                {
                    Width = (uint)((PageSize ?? PageSize.Default).WidthTwips()),
                    Height = (uint)((PageSize ?? PageSize.Default).HeightTwips()),
                },
                new W.PageMargin()
                {
                    // Notes: 
                    // - PageMargin uses uint for Left and Right margins, and int for top and bottom (enforced by Open XML SDK)
                    // - 0 is allowed for margins but not recommended
                    Left = (uint)((PageMargins ?? PageMargins.Default)).LeftTwips(),
                    Right = (uint)((PageMargins ?? PageMargins.Default)).RightTwips(),
                    Top = (int)((PageMargins ?? PageMargins.Default)).TopTwips(),
                    Bottom = (int)((PageMargins ?? PageMargins.Default)).BottomTwips(),
                }));
            
            if (targetDocument.CanSave)
                targetDocument.Save();         
        }
        catch(Exception)
        {
            throw;
        }
        finally
        {
            // Clear temp folder
            try 
            { 
                Directory.Delete(tempDir, true); 
            } 
            catch(Exception ex)
            { 
                #if DEBUG
                    Debug.WriteLine($"EPUB to DOCX: Unable to delete temp folder \"{tempDir}\". Details: {ex.Message}");
                #endif
                /* Write to console and ignore */ 
            }
        }  
    }
}
