using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToStringWriterBase<RtfStringWriter>
{
    internal override void ProcessSimpleField(SimpleField field, RtfStringWriter sb)
    {
        sb.WriteLine(@"{\field");
        if (field.Instruction?.Value != null)
        {
            // // TODO: FieldData
            // if (field.FieldData != null)
            // {
            // }

            // Open field instruction group
            sb.Write(@"{\*\fldinst {");

            // Write field instruction code. Fields may contain special characters such as '\' that need to be escaped.
            sb.WriteRtfEscaped(field.Instruction.Value);

            // Close field instruction group and open field result group.
            sb.Write(@"}}{\fldrslt {"); 

            // Process field result content
            base.ProcessSimpleField(field, sb); 

            // Close field result group and field destination.
            sb.WriteLine("}}}"); 
        }
    }

    internal override void ProcessFieldChar(FieldChar fieldChar, RtfStringWriter sb)
    {
        // Note: the content between the begin, separate and end parts is not processed here.
        if (fieldChar.FieldCharType != null)
        {
            if (fieldChar.FieldCharType == FieldCharValues.Begin)
            {
                sb.WriteLine(@"{\field");
                if (fieldChar.FieldLock != null && ((!fieldChar.FieldLock.HasValue) || fieldChar.FieldLock.Value))
                {
                    sb.Write("\\fldlock");
                }
                if (fieldChar.Dirty != null && ((!fieldChar.Dirty.HasValue) || fieldChar.Dirty.Value))
                {
                    sb.Write("\\flddirty");
                }

                if (fieldChar.FormFieldData != null)
                {
                    // The field is a form field
                    sb.Write(@"{{\*\formfield ");

                    if (fieldChar.FormFieldData.GetFirstChild<Enabled>().ToBool())
                    {
                        // Disabled
                        sb.Write(@"\ffprot1");
                    }
                    else
                    {
                        // Enabled (default)
                        sb.Write(@"\ffprot0");
                    }

                    if (fieldChar.FormFieldData.GetFirstChild<CalculateOnExit>().ToBool())
                    {
                        // Recalculate on exit if the CalculateOnExit element is present and not set to false
                        sb.Write(@"\ffrecalc1");
                    }
                    else
                    {
                        // Disable recalculate on exit
                        sb.Write(@"\ffrecalc0");
                    }

                    if (fieldChar.FormFieldData.GetFirstChild<TextInput>() is TextInput textInput)
                    {
                        sb.Write(@"\fftype0");

                        if (textInput.GetFirstChild<MaxLength>() is MaxLength maxLength &&
                           maxLength.Val != null)
                        {
                            sb.WriteWordWithValue("ffmaxlen", maxLength.Val.Value);
                        }
                        if (textInput.GetFirstChild<DefaultTextBoxFormFieldString>() is DefaultTextBoxFormFieldString defaultText &&
                            defaultText.Val != null && !string.IsNullOrEmpty(defaultText.Val.Value))
                        {
                            sb.Write($@"{{\*\ffdeftext {defaultText.Val.Value}}}");
                        }
                        if (textInput.GetFirstChild<TextBoxFormFieldType>() is TextBoxFormFieldType textBoxType &&
                           textBoxType.Val != null)
                        {
                            if (textBoxType.Val.Value == TextBoxFormFieldValues.Regular)
                            {
                                sb.Write(@"\fftypetxt0");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.Number)
                            {
                                sb.Write(@"\fftypetxt1");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.Date)
                            {
                                sb.Write(@"\fftypetxt2");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.CurrentDate)
                            {
                                sb.Write(@"\fftypetxt3");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.CurrentTime)
                            {
                                sb.Write(@"\fftypetxt4");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.Calculated)
                            {
                                sb.Write(@"\fftypetxt5");
                            }
                        }
                        if (textInput.GetFirstChild<Format>() is Format format &&
                            format.Val != null && !string.IsNullOrEmpty(format.Val.Value))
                        {
                            sb.Write($@"{{\*\ffformat {format.Val.Value}}}");
                        }
                    }
                    else if (fieldChar.FormFieldData.GetFirstChild<CheckBox>() is CheckBox checkBox)
                    {
                        sb.Write(@"\fftype1");

                        if (checkBox.GetFirstChild<FormFieldSize>() is FormFieldSize checkBoxSize &&
                            checkBoxSize.Val.ToLong() is long chkSize)
                        {
                            sb.WriteWordWithValue("ffhps", chkSize); // Check box size in half points
                        }

                        if (checkBox.GetFirstChild<AutomaticallySizeFormField>().ToBool())
                        {
                            sb.Write(@"\ffsize0"); // Auto size
                        }
                        else
                        {
                            sb.Write(@"\ffsize1"); // Exact size
                        }

                        if (checkBox.GetFirstChild<DefaultCheckBoxFormFieldState>() is DefaultCheckBoxFormFieldState defaultCheckBoxState &&
                            defaultCheckBoxState?.Val != null && defaultCheckBoxState.Val)
                        {
                            sb.Write(@"\ffdefres1"); // Checked by default
                        }
                        else
                        {
                            sb.Write(@"\ffdefres0"); // Unchecked by default
                        }

                        if (checkBox.GetFirstChild<Checked>().ToBool())
                        {
                            sb.Write(@"\ffres1"); // Checked if the Checked element is present and not false
                        }
                        else
                        {
                            sb.Write(@"\ffres0"); // Unchecked by default
                        }
                    }
                    else if (fieldChar.FormFieldData.GetFirstChild<DropDownListFormField>() is DropDownListFormField dropDownList)
                    {
                        sb.Write(@"\fftype2");

                        if (dropDownList.GetFirstChild<DefaultDropDownListItemIndex>() is DefaultDropDownListItemIndex defaultSelection &&
                            defaultSelection?.Val != null)
                        {
                            sb.WriteWordWithValue("ffdefres", defaultSelection.Val.Value); // Default selected index
                        }

                        if (dropDownList.GetFirstChild<DropDownListSelection>() is DropDownListSelection selection &&
                           selection?.Val != null)
                        {
                            sb.WriteWordWithValue("ffres", selection.Val.Value); // Current selected index
                        }

                        if (dropDownList.GetFirstChild<ListEntryFormField>() != null)
                        {
                            sb.Write(@"\ffhaslistbox1 ");
                            foreach (var listEntry in dropDownList.Elements<ListEntryFormField>())
                            {
                                if (listEntry.Val?.Value != null)
                                {
                                    sb.Write($@"{{\*\ffl {listEntry.Val.Value}}}");
                                }
                            }
                        }
                        else
                        {
                            sb.Write(@"\ffhaslistbox0");
                        }

                    }

                    if (fieldChar.GetFirstChild<StatusText>() is StatusText statusText)
                    {
                        if (statusText.Val?.Value != null)
                        {
                            sb.Write($@"\ffownstat1 {{\*\ffstattext {statusText.Val.Value}}}");
                        }
                        else
                        {
                            sb.Write(@"\ffownstat0");
                        }
                    }

                    if (fieldChar.GetFirstChild<HelpText>() is HelpText helpText)
                    {
                        if (helpText.Val?.Value != null)
                        {
                            sb.Write($@"\ffownhelp1 {{\*\ffhelptext {helpText.Val.Value}}}");
                        }
                        else
                        {
                            sb.Write(@"\ffownhelp0");
                        }
                    }

                    if (fieldChar.GetFirstChild<FormFieldName>() is FormFieldName name)
                    {
                        if (name.Val?.Value != null)
                        {
                            sb.Write($@"{{\*\ffname {name.Val.Value}}}");
                        }
                    }

                    if (fieldChar.GetFirstChild<EntryMacro>() is EntryMacro entryMacro)
                    {
                        if (entryMacro.Val?.Value != null)
                        {
                            sb.Write($@"{{\*\ffentrymcr {entryMacro.Val.Value}}}");
                        }
                    }

                    if (fieldChar.GetFirstChild<ExitMacro>() is ExitMacro exitMacro)
                    {
                        if (exitMacro.Val?.Value != null)
                        {
                            sb.Write($@"{{\*\ffexitmcr {exitMacro.Val.Value}}}");
                        }
                    }

                    sb.Write(@"}}");
                }

                // TODO: FieldData, NumberingChange
                //if (fieldChar.FieldData != null)
                //{
                //    // Custom field data (base64 binary value)
                //    sb.Append(@"{{\*\datafield ");
                //    // ...
                //    sb.Append(@"}}");
                //}

                // if (fieldChar.NumberingChange != null)
                // {
                // }

                sb.Write(@"{\*\fldinst {"); // Open field instruction group.
                //The last bracket is closed by the parent Run
            }
            else if (fieldChar.FieldCharType == FieldCharValues.Separate)
            {
                sb.Write(@"}}{\fldrslt {"); // Close field instruction group and open field result group.
                //The last bracket is closed by the parent Run.
            }
            else if (fieldChar.FieldCharType == FieldCharValues.End)
            {
                sb.WriteLine("}}}"); // Close field result and field destination.
            }
        }
    }

    internal override void ProcessFieldCode(FieldCode fieldCode, RtfStringWriter sb)
    {
        // Complex fields such as table of contents may contain special characters such as '\' that need to be escaped.
        sb.WriteRtfEscaped(fieldCode.InnerText);
    }
}
