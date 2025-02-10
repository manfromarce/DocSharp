using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    internal override void ProcessSdtBlock(SdtBlock sdtBlock, StringBuilder sb)
    {
        // TODO: sdtBlock.SdtProperties
        base.ProcessSdtBlock(sdtBlock, sb);
    }

    internal override void ProcessSdtRun(SdtRun sdtRun, StringBuilder sb)
    {
        // TODO: sdtRun.SdtProperties
        base.ProcessSdtRun(sdtRun, sb);
    }

    bool isInField = false;
    internal override void ProcessFieldChar(FieldChar fieldChar, StringBuilder sb)
    {
        if (fieldChar.FieldCharType != null)
        {
            if (fieldChar.FieldCharType == FieldCharValues.Begin)
            {
                isInField = true;
                sb.Append(@"{\field");
                if (fieldChar.FieldLock != null && ((!fieldChar.FieldLock.HasValue) || fieldChar.FieldLock.Value))
                {
                    sb.Append("\\fldlock");
                }
                sb.Append(@"{\*\fldinst {");
            }
            else if (fieldChar.FieldCharType == FieldCharValues.Separate)
            {
                sb.Append(@"}}{\fldrslt ");
            }
            else if (fieldChar.FieldCharType == FieldCharValues.End)
            {
                sb.Append("}}");
                isInField = false;
            }
        }
    }

    internal override void ProcessFieldCode(FieldCode fieldCode, StringBuilder sb)
    {
        if (isInField)
        {
            sb.Append(fieldCode.InnerText);
        }
    }

    internal override void ProcessSimpleField(SimpleField simpleField, StringBuilder sb)
    {
        //sb.Append("{\\field");
        //if (simpleField.Instruction != null)
        //{
        //}
        //if (simpleField.FieldData != null)
        //{
        //}
        //if (simpleField.FieldLock != null && ((!simpleField.FieldLock.HasValue) || simpleField.FieldLock.Value)) 
        //{
        //    sb.Append("\\fldlock");
        //}
        base.ProcessSimpleField(simpleField, sb);
        //sb.Append("}");
    }
}
