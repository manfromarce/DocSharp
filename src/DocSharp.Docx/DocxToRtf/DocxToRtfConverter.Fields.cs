using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
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

                //if (fieldChar.FieldData != null)
                //{
                //    // Custom field data (base64 binary value)
                //    sb.Append(@"{{\*\datafield ");
                //    // ...
                //    sb.Append(@"}}");
                //}

                if (fieldChar.FormFieldData != null)
                {
                    // The field is a form field
                    sb.Write(@"{{\*\formfield ");

                    if (fieldChar.FormFieldData.GetFirstChild<Enabled>() is Enabled enabled && 
                        enabled.Val != null && !enabled.Val)
                    {
                        // Disabled
                        sb.Write(@"\ffprot1");
                    }
                    else
                    {
                        // Enabled (default)
                        sb.Write(@"\ffprot0");
                    }

                    if (fieldChar.FormFieldData.GetFirstChild<CalculateOnExit>() is CalculateOnExit calcOnExit &&
                       calcOnExit != null && (calcOnExit.Val == null || calcOnExit.Val))
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
                            sb.Write($@"\ffmaxlen{maxLength.Val}");
                        }
                        if (textInput.GetFirstChild<DefaultTextBoxFormFieldString>() is DefaultTextBoxFormFieldString defaultText &&
                           defaultText.Val != null)
                        {
                            sb.Write($@"{{\*\ffdeftext {defaultText.Val}}}");
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
                           format.Val != null)
                        {
                            sb.Write($@"{{\*\ffformat {format.Val}}}");
                        }
                    }
                    else if (fieldChar.FormFieldData.GetFirstChild<CheckBox>() is CheckBox checkBox)
                    {
                        sb.Write(@"\fftype1");

                        if (checkBox.GetFirstChild<FormFieldSize>() is FormFieldSize checkBoxSize &&
                            checkBoxSize.Val != null)
                        {
                            sb.Write($@"\ffhps{checkBoxSize.Val}"); // Check box size in half points
                        }

                        if (checkBox.GetFirstChild<AutomaticallySizeFormField>() is AutomaticallySizeFormField autoSize &&
                            autoSize != null && (autoSize.Val == null || autoSize.Val))
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

                        if (checkBox.GetFirstChild<Checked>() is Checked @checked &&
                            @checked != null && (@checked.Val == null || @checked.Val))
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
                            sb.Write(@$"\ffdefres{defaultSelection?.Val}"); // Default selected index
                        }

                        if (dropDownList.GetFirstChild<DropDownListSelection>() is DropDownListSelection selection &&
                           selection?.Val != null)
                        {
                            sb.Write(@$"\ffres{selection?.Val}"); // Current selected index
                        }

                        if (dropDownList.GetFirstChild<ListEntryFormField>() != null)
                        {
                            sb.Write(@"\ffhaslistbox1 ");
                            foreach (var listEntry in dropDownList.Elements<ListEntryFormField>())
                            {
                                if (listEntry.Val != null)
                                {
                                    sb.Write($@"{{\*\ffl {listEntry.Val}}}");
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
                        if (statusText.Val != null)
                        {
                            sb.Write($@"\ffownstat1 {{\*\ffstattext {statusText.Val}}}");
                        }
                        else
                        {
                            sb.Write(@"\ffownstat0");
                        }
                    }

                    if (fieldChar.GetFirstChild<HelpText>() is HelpText helpText)
                    {
                        if (helpText.Val != null)
                        {
                            sb.Write($@"\ffownhelp1 {{\*\ffhelptext {helpText.Val}}}");
                        }
                        else
                        {
                            sb.Write(@"\ffownhelp0");
                        }
                    }

                    if (fieldChar.GetFirstChild<FormFieldName>() is FormFieldName name)
                    {
                        if (name.Val != null)
                        {
                            sb.Write($@"{{\*\ffname {name.Val}}}");
                        }
                    }

                    if (fieldChar.GetFirstChild<EntryMacro>() is EntryMacro entryMacro)
                    {
                        if (entryMacro.Val != null)
                        {
                            sb.Write($@"{{\*\ffentrymcr {entryMacro.Val}}}");
                        }
                    }

                    if (fieldChar.GetFirstChild<ExitMacro>() is ExitMacro exitMacro)
                    {
                        if (exitMacro.Val != null)
                        {
                            sb.Write($@"{{\*\ffexitmcr {exitMacro.Val}}}");
                        }
                    }

                    sb.Write(@"}}");
                }

                sb.Write(@"{\*\fldinst {"); // Open field instruction group.
                //The last bracket is closed by the parent Run
            }
            else if (fieldChar.FieldCharType == FieldCharValues.Separate)
            {
                sb.Write(@"}}{\fldrslt {"); // Close field instruction and open field result group.
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
