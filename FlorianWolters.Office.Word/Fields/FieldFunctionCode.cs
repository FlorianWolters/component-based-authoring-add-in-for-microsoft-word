//------------------------------------------------------------------------------
// <copyright file="FieldFunctionCode.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System.Text.RegularExpressions;
    using FlorianWolters.Office.Word.Fields.Switches;

    /// <summary>
    /// The class <see cref="FieldFunctionCode"/> allows to manipulate the code of a Microsoft Word field.
    /// </summary>
    public class FieldFunctionCode
    {
        /// <summary>
        /// The delimiter used to specify the format of a field.
        /// </summary>
        public const string FormatFieldSwitch = @"\*";

        /// <summary>
        /// The position of the <see cref="FormatFieldSwitch"/> in the code of the field.
        /// </summary>
        private int formatFieldSwitchPosition;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFunctionCode"/> class with the specified field code.
        /// </summary>
        /// <param name="input">The field code.</param>
        public FieldFunctionCode(string input)
        {
            this.formatFieldSwitchPosition = input.IndexOf(FormatFieldSwitch);

            if (this.ContainsFormatSwitch())
            {
                this.Function = this.NormalizeFunction(input);
                this.Format = this.NormalizeFormat(input);
            }
            else
            {
                this.Function = input.Trim();
                this.Format = string.Empty;
            }
        }

        /// <summary>
        /// Gets the function part of this <see cref="FieldFunctionCode"/>.
        /// </summary>
        public string Function { get; private set; }

        /// <summary>
        /// Gets the format part of this <see cref="FieldFunctionCode"/>.
        /// </summary>
        public string Format { get; private set; }

        /// <summary>
        /// Returns a string representation of this <see cref="FieldFunctionCode"/>.
        /// </summary>
        /// <returns>The string representation.</returns>
        public override string ToString()
        {
            string result = this.Function;

            if (string.Empty != this.Format)
            {
                result += " " + this.Format;
            }

            return result;
        }

        /// <summary>
        /// Checks whether this <see cref="FieldFunctionCode"/> contains the specified field format switch.
        /// </summary>
        /// <param name="fieldFormatSwitch">The format switch to search for.</param>
        /// <returns>
        /// <c>true</c> whether this <see cref="FieldFunctionCode"/> contains the specified field format switch;
        /// <c>false</c> otherwise.
        /// </returns>
        public bool ContainsFormatSwitch(FieldFormatSwitches fieldFormatSwitch)
        {
            return this.Format.Contains(fieldFormatSwitch.ToString());
        }

        /// <summary>
        /// Checks whether this <see cref="FieldFunctionCode"/> contains at least one field format switch.
        /// </summary>
        /// <returns>
        /// <c>true</c> whether this <see cref="FieldFunctionCode"/> contains at least one  field format switch;
        /// <c>false</c> otherwise.
        /// </returns>
        public bool ContainsFormatSwitch()
        {
            return -1 != this.formatFieldSwitchPosition;
        }

        /// <summary>
        /// Normalizes the specified function part of a field code. 
        /// </summary>
        /// <param name="input">The function part.</param>
        /// <returns>The normalized function part.</returns>
        private string NormalizeFunction(string input)
        {
            return input.Substring(0, this.formatFieldSwitchPosition).Trim();
        }

        /// <summary>
        /// Normalizes the specified format part of a field code. 
        /// </summary>
        /// <param name="input">The format part.</param>
        /// <returns>The normalized format part.</returns>
        private string NormalizeFormat(string input)
        {
            string result = input.Substring(this.formatFieldSwitchPosition).Trim();

            string pattern = "(" + Regex.Escape(FormatFieldSwitch) + @")(\s+)(.)";
            result = Regex.Replace(result, pattern, "$1$3");
            
            return Regex.Replace(result, @"\s+", " ").ToUpper();
        }
    }
}
