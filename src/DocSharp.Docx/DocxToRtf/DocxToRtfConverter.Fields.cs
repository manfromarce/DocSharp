using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{

    internal override void ProcessFieldChar(FieldChar fieldChar, StringBuilder sb)
    {
        // Note: the content between the begin, separate and end parts is not processed here.
        if (fieldChar.FieldCharType != null)
        {
            if (fieldChar.FieldCharType == FieldCharValues.Begin)
            {
                sb.AppendLineCrLf(@"{\field");
                if (fieldChar.FieldLock != null && ((!fieldChar.FieldLock.HasValue) || fieldChar.FieldLock.Value))
                {
                    sb.Append("\\fldlock");
                }
                if (fieldChar.Dirty != null && ((!fieldChar.Dirty.HasValue) || fieldChar.Dirty.Value))
                {
                    sb.Append("\\flddirty");
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
                    sb.Append(@"{{\*\formfield ");

                    if (fieldChar.FormFieldData.GetFirstChild<Enabled>() is Enabled enabled && 
                        enabled.Val != null && !enabled.Val)
                    {
                        // Disabled
                        sb.Append(@"\ffprot1");
                    }
                    else
                    {
                        // Enabled (default)
                        sb.Append(@"\ffprot0");
                    }

                    if (fieldChar.FormFieldData.GetFirstChild<CalculateOnExit>() is CalculateOnExit calcOnExit &&
                       calcOnExit != null && (calcOnExit.Val == null || calcOnExit.Val))
                    {
                        // Recalculate on exit if the CalculateOnExit element is present and not set to false
                        sb.Append(@"\ffrecalc1");
                    }
                    else
                    {
                        // Disable recalculate on exit
                        sb.Append(@"\ffrecalc0");
                    }

                    if (fieldChar.FormFieldData.GetFirstChild<TextInput>() is TextInput textInput)
                    {
                        sb.Append(@"\fftype0");

                        if (textInput.GetFirstChild<MaxLength>() is MaxLength maxLength &&
                           maxLength.Val != null)
                        {
                            sb.Append($@"\ffmaxlen{maxLength.Val}");
                        }
                        if (textInput.GetFirstChild<DefaultTextBoxFormFieldString>() is DefaultTextBoxFormFieldString defaultText &&
                           defaultText.Val != null)
                        {
                            sb.Append($@"{{\*\ffdeftext {defaultText.Val}}}");
                        }
                        if (textInput.GetFirstChild<TextBoxFormFieldType>() is TextBoxFormFieldType textBoxType &&
                           textBoxType.Val != null)
                        {
                            if (textBoxType.Val.Value == TextBoxFormFieldValues.Regular)
                            {
                                sb.Append(@"\fftypetxt0");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.Number)
                            {
                                sb.Append(@"\fftypetxt1");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.Date)
                            {
                                sb.Append(@"\fftypetxt2");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.CurrentDate)
                            {
                                sb.Append(@"\fftypetxt3");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.CurrentTime)
                            {
                                sb.Append(@"\fftypetxt4");
                            }
                            else if (textBoxType.Val.Value == TextBoxFormFieldValues.Calculated)
                            {
                                sb.Append(@"\fftypetxt5");
                            }
                        }
                        if (textInput.GetFirstChild<Format>() is Format format &&
                           format.Val != null)
                        {
                            sb.Append($@"{{\*\ffformat {format.Val}}}");
                        }
                    }
                    else if (fieldChar.FormFieldData.GetFirstChild<CheckBox>() is CheckBox checkBox)
                    {
                        sb.Append(@"\fftype1");

                        if (checkBox.GetFirstChild<FormFieldSize>() is FormFieldSize checkBoxSize &&
                            checkBoxSize.Val != null)
                        {
                            sb.Append($@"\ffhps{checkBoxSize.Val}"); // Check box size in half points
                        }

                        if (checkBox.GetFirstChild<AutomaticallySizeFormField>() is AutomaticallySizeFormField autoSize &&
                            autoSize != null && (autoSize.Val == null || autoSize.Val))
                        {
                            sb.Append(@"\ffsize0"); // Auto size
                        }
                        else
                        {
                            sb.Append(@"\ffsize1"); // Exact size
                        }

                        if (checkBox.GetFirstChild<DefaultCheckBoxFormFieldState>() is DefaultCheckBoxFormFieldState defaultCheckBoxState &&
                            defaultCheckBoxState?.Val != null && defaultCheckBoxState.Val)
                        {
                            sb.Append(@"\ffdefres1"); // Checked by default
                        }
                        else
                        {
                            sb.Append(@"\ffdefres0"); // Unchecked by default
                        }

                        if (checkBox.GetFirstChild<Checked>() is Checked @checked &&
                            @checked != null && (@checked.Val == null || @checked.Val))
                        {
                            sb.Append(@"\ffres1"); // Checked if the Checked element is present and not false
                        }
                        else
                        {
                            sb.Append(@"\ffres0"); // Unchecked by default
                        }
                    }
                    else if (fieldChar.FormFieldData.GetFirstChild<DropDownListFormField>() is DropDownListFormField dropDownList)
                    {
                        sb.Append(@"\fftype2");

                        if (dropDownList.GetFirstChild<DefaultDropDownListItemIndex>() is DefaultDropDownListItemIndex defaultSelection &&
                            defaultSelection?.Val != null)
                        {
                            sb.Append(@$"\ffdefres{defaultSelection?.Val}"); // Default selected index
                        }

                        if (dropDownList.GetFirstChild<DropDownListSelection>() is DropDownListSelection selection &&
                           selection?.Val != null)
                        {
                            sb.Append(@$"\ffres{selection?.Val}"); // Current selected index
                        }

                        if (dropDownList.GetFirstChild<ListEntryFormField>() != null)
                        {
                            sb.Append(@"\ffhaslistbox1 ");
                            foreach (var listEntry in dropDownList.Elements<ListEntryFormField>())
                            {
                                if (listEntry.Val != null)
                                {
                                    sb.Append($@"{{\*\ffl {listEntry.Val}}}");
                                }
                            }
                        }
                        else
                        {
                            sb.Append(@"\ffhaslistbox0");
                        }

                    }

                    if (fieldChar.GetFirstChild<StatusText>() is StatusText statusText)
                    {
                        if (statusText.Val != null)
                        {
                            sb.Append($@"\ffownstat1 {{\*\ffstattext {statusText.Val}}}");
                        }
                        else
                        {
                            sb.Append(@"\ffownstat0");
                        }
                    }

                    if (fieldChar.GetFirstChild<HelpText>() is HelpText helpText)
                    {
                        if (helpText.Val != null)
                        {
                            sb.Append($@"\ffownhelp1 {{\*\ffhelptext {helpText.Val}}}");
                        }
                        else
                        {
                            sb.Append(@"\ffownhelp0");
                        }
                    }

                    if (fieldChar.GetFirstChild<FormFieldName>() is FormFieldName name)
                    {
                        if (name.Val != null)
                        {
                            sb.Append($@"{{\*\ffname {name.Val}}}");
                        }
                    }

                    if (fieldChar.GetFirstChild<EntryMacro>() is EntryMacro entryMacro)
                    {
                        if (entryMacro.Val != null)
                        {
                            sb.Append($@"{{\*\ffentrymcr {entryMacro.Val}}}");
                        }
                    }

                    if (fieldChar.GetFirstChild<ExitMacro>() is ExitMacro exitMacro)
                    {
                        if (exitMacro.Val != null)
                        {
                            sb.Append($@"{{\*\ffexitmcr {exitMacro.Val}}}");
                        }
                    }

                    sb.Append(@"}}");
                }

                sb.Append(@"{\*\fldinst {"); // Open field instruction group.
                //The last bracket is closed by the parent Run
            }
            else if (fieldChar.FieldCharType == FieldCharValues.Separate)
            {
                sb.Append(@"}}{\fldrslt {"); // Close field instruction and open field result group.
                //The last bracket is closed by the parent Run.
            }
            else if (fieldChar.FieldCharType == FieldCharValues.End)
            {
                sb.AppendLineCrLf("}}}"); // Close field result and field destination.
            }
        }
    }

    internal override void ProcessFieldCode(FieldCode fieldCode, StringBuilder sb)
    {
        // Complex fields such as table of contents may contain special characters such as '\' that need to be escaped.
        sb.AppendRtfEscaped(fieldCode.InnerText);
    }
}
