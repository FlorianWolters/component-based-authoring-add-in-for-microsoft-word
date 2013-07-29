//------------------------------------------------------------------------------
// <copyright file="FieldFormatSwitch.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.Switches
{
    using System.Text.RegularExpressions;
    using Word = Microsoft.Office.Interop.Word;

    // http://office.microsoft.com/word-help/format-field-switch-HP005186222.aspx
    public class FieldFormatSwitch
    {
        private readonly string switchName;
        private readonly string searchPattern;

        public FieldFormatSwitch(string switchName)
        {
            this.switchName = switchName;
            this.searchPattern = Regex.Escape(FieldFunctionCode.FormatFieldSwitch) + @"\s*" + switchName;
        }

        public void AddToField(Word.Field field)
        {
            field.Code.Text = this.AddToString(field.Code.Text);
        }

        public string AddToString(string fieldCodeText)
        {
            string result = fieldCodeText.Trim();

            if (!Regex.IsMatch(result, this.searchPattern, RegexOptions.IgnoreCase))
            {
                result += " " + FieldFunctionCode.FormatFieldSwitch + this.switchName;
            }

            return new FieldFunctionCode(result).ToString();
        }

        public void RemoveFromField(Word.Field field)
        {
            field.Code.Text = this.RemoveFromString(field.Code.Text);
        }

        public string RemoveFromString(string fieldCodeText)
        {
            string result = fieldCodeText.Trim();

            if (Regex.IsMatch(fieldCodeText, this.searchPattern, RegexOptions.IgnoreCase))
            {
                result = Regex.Replace(fieldCodeText, this.searchPattern, string.Empty, RegexOptions.IgnoreCase);
            }

            FieldFunctionCode parser = new FieldFunctionCode(result);
            result = parser.ToString();

            return result;
        }
    }
}
