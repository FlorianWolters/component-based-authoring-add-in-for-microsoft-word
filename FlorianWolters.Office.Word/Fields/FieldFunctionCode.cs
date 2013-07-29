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

    public class FieldFunctionCode
    {
        public const string FormatFieldSwitch = @"\*";
        private int formFieldSwitchPosition;

        public FieldFunctionCode(string input)
        {
            this.SetFirstFormatFieldSwitchPosition(input);

            if (this.ContainsFormatFieldSwitch())
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

        public override string ToString()
        {
            string result = this.Function;

            if (string.Empty != this.Format)
            {
                result += " " + this.Format;
            }

            return result;
        }

        public bool ContainsFormatSwitch(FieldFormatSwitches fieldFormatSwitch)
        {
            return this.Format.Contains(fieldFormatSwitch.ToString());
        }

        private bool ContainsFormatFieldSwitch()
        {
            return -1 != this.formFieldSwitchPosition;
        }

        private void SetFirstFormatFieldSwitchPosition(string input)
        {
            this.formFieldSwitchPosition = input.IndexOf(FormatFieldSwitch);
        }

        private string NormalizeFunction(string input)
        {
            return input.Substring(0, this.formFieldSwitchPosition).Trim();
        }

        private string NormalizeFormat(string input)
        {
            string result = input.Substring(this.formFieldSwitchPosition).Trim();

            string pattern = "(" + Regex.Escape(FormatFieldSwitch) + @")(\s+)(.)";
            result = Regex.Replace(result, pattern, "$1$3");
            
            return Regex.Replace(result, @"\s+", " ").ToUpper();
        }
    }
}
