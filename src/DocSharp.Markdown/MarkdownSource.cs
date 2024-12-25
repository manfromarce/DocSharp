using System.IO;
using Markdig;
using Markdig.Syntax;

namespace DocSharp.Markdown;

/// <summary>
/// Represents a source of Markdown content. It implicitly converts from a 
/// <see cref="Markdig.Syntax.MarkdownDocument"/> or <see cref="Stream" />,
/// or can be created from file path or <see cref="string"/> using  
/// the FromFile and FromMarkdownString static functions.
/// </summary>
public class MarkdownSource
{
    /// <summary>
    /// The Markdown Document associated to this source.
    /// </summary>
    public MarkdownDocument Document { get; }

    /// <summary>
    /// Create a Markdown source from a document
    /// </summary>
    /// <param name="document">The <see cref="Markdig.Syntax.MarkdownDocument"/> to use</param>
    public MarkdownSource(MarkdownDocument document)
    {
        Document = document;
    }

    /// <summary>
    /// Create a Markdown source from a file path
    /// </summary>
    /// <param name="filePath">The file path to use</param>
    public static MarkdownSource FromFile(string filePath)
    {
        return FromMarkdownString(File.ReadAllText(filePath));
    }

    /// <summary>
    /// Create a Markdown source from a stream
    /// </summary>
    /// <param name="stream">The stream to use</param>
    public static MarkdownSource FromStream(Stream stream)
    {
        using (var streamReader = new StreamReader(stream))
        {
            string markdown = streamReader.ReadToEnd();
            return MarkdownSource.FromMarkdownString(markdown);
        }
    }

    /// <summary>
    /// Create a Markdown source from a Markdown string
    /// </summary>
    /// <param name="markdown">The Markdown content as string</param>
    public static MarkdownSource FromMarkdownString(string markdown)
    {
        var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions()
                                                    .UseEmojiAndSmiley()
                                                    .Build();

        return Markdig.Markdown.Parse(markdown, pipeline);
    }

    /// <summary>
    /// Implicitly convert a <see cref="MarkdownDocument"/> to a <see cref="MarkdownSource"/>
    /// </summary>
    /// <param name="document">Markdown document to use</param>
    public static implicit operator MarkdownSource(MarkdownDocument document)
    {
        return new MarkdownSource(document);
    }

    /// <summary>
    /// Implicitly convert a <see cref="Stream"/> containing Markdown content to a <see cref="MarkdownSource"/>
    /// </summary>
    /// <param name="stream">Markdown content stream</param>
    public static implicit operator MarkdownSource(Stream stream)
    {
        return MarkdownSource.FromStream(stream);
    }
}

