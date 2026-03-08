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
    private void CreateSimpleField(string instr, string currentValue)
    {
        AddRun().Append(new SimpleField(new Run(new Text(currentValue)))
        {
            Instruction = instr
        });       

        // Ensure that the following content is added to a new run
        currentRun = null;
    }
    
    private void CreateField(string instrText, string currentValue)
    {
        // Part 1 - Begin
        AddRun().Append(new FieldChar()
        {
            FieldCharType = FieldCharValues.Begin
        });

        // Part 2 - InstrText
        AddRun().Append(new FieldCode(instrText ?? string.Empty));

        // Part 3 - Separate
        AddRun().Append(new FieldChar()
        {
            FieldCharType = FieldCharValues.Separate
        });

        // Part 4 - Current value
        AddRun().Append(new Text(currentValue ?? string.Empty)
        {
            Space = SpaceProcessingModeValues.Preserve
        });

        // Part 5 - End
        AddRun().Append(new FieldChar()
        {
            FieldCharType = FieldCharValues.End
        });

        // Ensure that the following content is added to a new run
        currentRun = null;
    }
}