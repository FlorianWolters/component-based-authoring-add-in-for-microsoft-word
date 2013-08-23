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

    /// <summary>
    /// The class <see cref="FieldFormatSwitch"/> wraps a
    /// <a href="http://office.microsoft.com/en-us/word-help/format-field-switch-HP005186222.aspx">format field switch</a>
    /// into an object.
    /// </summary>
    public class FieldFormatSwitch
    {
        /// <summary>
        /// The name of this <see cref="FieldFormatSwitch"/>.
        /// </summary>
        private readonly string switchName;

        /// <summary>
        /// The search pattern to find a switch in the field code.
        /// </summary>
        private readonly string searchPattern;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFormatSwitch"/> class.
        /// </summary>
        /// <param name="switchName">The name of the format field switch.</param>
        public FieldFormatSwitch(string switchName)
        {
            this.switchName = switchName;
            this.searchPattern = Regex.Escape(FieldFunctionCode.FormatFieldSwitch) + @"\s*" + switchName;
        }

        /// <summary>
        /// Adds this <see cref="FieldFormatSwitch"/> to the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to manipulate.</param>
        public void AddToField(Word.Field field)
        {
            field.Code.Text = this.AddToString(field.Code.Text);
        }

        /// <summary>
        /// Adds this <see cref="FieldFormatSwitch"/> to the specified <c>string</c>.
        /// </summary>
        /// <param name="fieldCodeText">The code of a <see cref="Word.Field"/>.</param>
        /// <returns>The manipulated string.</returns>
        public string AddToString(string fieldCodeText)
        {
            string result = fieldCodeText.Trim();

            if (!Regex.IsMatch(result, this.searchPattern, RegexOptions.IgnoreCase))
            {
                result += " " + FieldFunctionCode.FormatFieldSwitch + this.switchName;
            }

            return new FieldFunctionCode(result).ToString();
        }

        /// <summary>
        /// Removes this <see cref="FieldFormatSwitch"/> from the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to manipulate.</param>
        public void RemoveFromField(Word.Field field)
        {
            field.Code.Text = this.RemoveFromString(field.Code.Text);
        }

        /// <summary>
        /// Removed this <see cref="FieldFormatSwitch"/> from the specified <c>string</c>.
        /// </summary>
        /// <param name="fieldCodeText">The code of a <see cref="Word.Field"/>.</param>
        /// <returns>The manipulated string.</returns>
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
