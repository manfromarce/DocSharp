using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override void ProcessCommentStart(CommentRangeStart commentStart, RtfStringWriter sb)
    {
        if (commentStart.Id?.Value != null)
        {
            sb.Write(@$"{{\*\atrfstart {commentStart.Id.Value}}}");
        }
    }

    internal override void ProcessCommentEnd(CommentRangeEnd commentEnd, RtfStringWriter sb)
    {
        if (commentEnd.Id?.Value != null)
        {
            sb.Write(@$"{{\*\atrfend {commentEnd.Id.Value}}}");
        }
    }

    internal override void ProcessCommentReference(CommentReference commentRef, RtfStringWriter sb)
    {
        var root = commentRef.GetMainDocumentPart();
        if (commentRef.Id?.Value != null && 
            root?.WordprocessingCommentsPart?.Comments is Comments comments &&
            comments.Elements<Comment>().Where(c => c.Id?.Value != null && c.Id.Value == commentRef.Id.Value).FirstOrDefault() is Comment comment)
        {
            if (comment.Initials?.Value != null)
            {
                sb.Write(@$"{{\*\atnid {comment.Initials.Value}}}");
            }
            if (comment.Author?.Value != null)
            {
                sb.Write(@$"{{\*\atnauthor {comment.Author.Value}}}");
            }

            sb.Write(@"\chatn {\*\annotation"); // Write annotation destination
            sb.Write(@$"{{\*\atnref {commentRef.Id.Value}}}");
            if (comment.Date != null)
            {
                sb.Write(@"{\*\atndate ");
                sb.WriteRtfDate(comment.Date.Value);
                sb.Write(@"}");
            }

            // Process comment content
            foreach (var element in comment.Elements())
            {
                ProcessBodyElement(element, sb);
            }
            
            sb.WriteLine(@"}"); // Close annotation destination
        }
    }

    internal override void ProcessAnnotationReference(AnnotationReferenceMark annotationRef, RtfStringWriter sb)
    {
        sb.Write(@"\chatn");
    }
}