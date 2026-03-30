namespace DocSharp.Docx;

/// <summary>
/// Maps DOCX style IDs to known style types (e.g. Heading 1, Quote, etc.) (see also <see cref="StyleType"/>).
/// </summary>
public interface IStyleNamingResolver
{
    /// <summary>
    /// Tries to determine the <see cref="StyleType"/> for the given <paramref name="styleId"/>.
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <param name="styleType">The corresponding <see cref="StyleType"/> if found.</param>
    /// <returns><see langword="true"/> if the style type could be determined, otherwise <see langword="false"/>.</returns>
    bool TryGetStyleType(string? styleId, out StyleType styleType);
}

/// <summary>
/// Specifies common DOCX styles.
/// </summary>
public enum StyleType
{
    /// <summary>
    /// Header 1 / Title
    /// </summary>
    Header1 = 0,

    /// <summary>
    /// Header 2 / Subtitle
    /// </summary>
    Header2 = 1,

    /// <summary>
    /// Header 3
    /// </summary>
    Header3 = 2,

    /// <summary>
    /// Header 4
    /// </summary>
    Header4 = 3,

    /// <summary>
    /// Header 5
    /// </summary>
    Header5 = 4,

    /// <summary>
    /// Header 6
    /// </summary>
    Header6 = 5,

    /// <summary>
    /// Quote
    /// </summary>
    Quote = 6,

    /// <summary>
    /// Intense Quote
    /// </summary>
    IntenseQuote = 7,

    /// <summary>
    /// HTML Preformatted
    /// </summary>
    HtmlPreformatted = 8,
}
