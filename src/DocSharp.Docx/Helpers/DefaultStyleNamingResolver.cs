using System;

namespace DocSharp.Docx;

/// <summary>
/// Provides a default implementation of the IStyleNamingResolver interface for mapping style identifiers to their
/// corresponding style types, such as headers and quotes, in document processing scenarios.
/// </summary>
/// <remarks>Can be overridden to provide custom style naming logic.</remarks>
public class DefaultStyleNamingResolver : IStyleNamingResolver
{
    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a 1st level header style / title, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHeader1Style(string styleId)
    {
        return styleId.Equals("Heading1", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Heading 1", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Title", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a 2nd level header style / subtitle, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHeader2Style(string styleId)
    {
        return styleId.Equals("Heading2", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Heading 2", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a 3rd level header style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHeader3Style(string styleId)
    {
        return styleId.Equals("Heading3", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Heading 3", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a 4th level header style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHeader4Style(string styleId)
    {
        return styleId.Equals("Heading4", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Heading 4", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a 5th level header style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHeader5Style(string styleId)
    {
        return styleId.Equals("Heading5", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Heading 5", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a 6th level header style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHeader6Style(string styleId)
    {
        return styleId.Equals("Heading6", StringComparison.OrdinalIgnoreCase) ||
               styleId.Equals("Heading 6", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a quote style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsQuoteStyle(string styleId)
    {
        return styleId.Equals("Quote", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to an intense quote style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsIntenseQuoteStyle(string styleId)
    {
        return styleId.Equals("Intense Quote", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Returns <see langword="true"/> if the given <paramref name="styleId"/> corresponds to a html preformatted style, otherwise <see langword="false"/>.<br/>
    /// </summary>
    /// <param name="styleId">The style ID to check.</param>
    /// <remarks>This style is created by Microsoft Word when an HTML file is saved as DOCX</remarks>
    /// <returns><see langword="true"/> if the style matches, otherwise <see langword="false"/>.</returns>
    public virtual bool IsHtmlPreformattedStyle(string styleId)
    {
        return styleId.Equals("Html Preformatted", StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    public virtual bool TryGetStyleType(string? styleId, out StyleType styleType)
    {
        if (styleId == null || string.IsNullOrEmpty(styleId))
        {
            styleType = default;
            return false;
        }
        if (IsHeader1Style(styleId))
        {
            styleType = StyleType.Header1;
            return true;
        }
        if (IsHeader2Style(styleId))
        {
            styleType = StyleType.Header2;
            return true;
        }
        if (IsHeader3Style(styleId))
        {
            styleType = StyleType.Header3;
            return true;
        }
        if (IsHeader4Style(styleId))
        {
            styleType = StyleType.Header4;
            return true;
        }
        if (IsHeader5Style(styleId))
        {
            styleType = StyleType.Header5;
            return true;
        }
        if (IsHeader6Style(styleId))
        {
            styleType = StyleType.Header6;
            return true;
        }
        if (IsQuoteStyle(styleId))
        {
            styleType = StyleType.Quote;
            return true;
        }
        if (IsIntenseQuoteStyle(styleId))
        {
            styleType = StyleType.IntenseQuote;
            return true;
        }
        if (IsHtmlPreformattedStyle(styleId))
        {
            styleType = StyleType.HtmlPreformatted;
            return true;
        }

        styleType = default;
        return false;
    }
}
