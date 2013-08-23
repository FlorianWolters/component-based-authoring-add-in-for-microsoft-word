//------------------------------------------------------------------------------
// <copyright file="FieldFormatSwitches.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.Switches
{
    using FlorianWolters.Office.Word.Fields;

    /// <summary>
    /// The class <see cref="FieldFormatSwitches"/> enumerates format field switches for a field in Microsoft Word.
    /// </summary>
    public sealed class FieldFormatSwitches
    {
        /// <summary>
        /// Capitalizes the first letter of each word.
        /// </summary>
        public static readonly FieldFormatSwitches Caps = new FieldFormatSwitches(1, "CAPS");

        /// <summary>
        /// Capitalizes the first letter of the first word.
        /// </summary>
        public static readonly FieldFormatSwitches FirstCap = new FieldFormatSwitches(2, "FIRSTCAP");

        /// <summary>
        /// Capitalizes all letters.
        /// </summary>
        public static readonly FieldFormatSwitches Upper = new FieldFormatSwitches(3, "UPPER");

        /// <summary>
        /// Converts all letters to lowercase letters.
        /// </summary>
        public static readonly FieldFormatSwitches Lower = new FieldFormatSwitches(4, "LOWER");

        /// <summary>
        /// Displays results as alphabetic characters.
        /// </summary>
        public static readonly FieldFormatSwitches Alphabetic = new FieldFormatSwitches(5, "ALPHABETIC");

        /// <summary>
        /// Displays results as Arabic cardinal numerals.
        /// </summary>
        public static readonly FieldFormatSwitches Arabic = new FieldFormatSwitches(6, "ARABIC");

        /// <summary>
        /// Displays results as cardinal text.
        /// </summary>
        public static readonly FieldFormatSwitches CardText = new FieldFormatSwitches(7, "CARDTEXT");

        /// <summary>
        /// Displays results as cardinal text. Microsoft Word inserts "and" at the decimal place and displays the first
        /// two decimals (rounded) as Arabic numerators over 100.
        /// </summary>
        public static readonly FieldFormatSwitches DollarText = new FieldFormatSwitches(8, "DOLLARTEXT");

        /// <summary>
        /// Displays results as hexadecimal numbers.
        /// </summary>
        public static readonly FieldFormatSwitches Hex = new FieldFormatSwitches(9, "HEX");

        /// <summary>
        /// Displays results as ordinal text.
        /// </summary>
        public static readonly FieldFormatSwitches OrdText = new FieldFormatSwitches(10, "ORDTEXT");

        /// <summary>
        /// Displays results as ordinal Arabic numerals.
        /// </summary>
        public static readonly FieldFormatSwitches Ordinal = new FieldFormatSwitches(11, "ORDINAL");

        /// <summary>
        /// Displays results as Roman numerals. 
        /// </summary>
        public static readonly FieldFormatSwitches Roman = new FieldFormatSwitches(12, "ROMAN");

        /// <summary>
        /// Applies the formatting of the first letter of the field type to the entire result.
        /// </summary>
        public static readonly FieldFormatSwitches CharFormat = new FieldFormatSwitches(13, "CHARFORMAT");

        /// <summary>
        /// Applies the formatting of the previous result to the new result.
        /// </summary>
        public static readonly FieldFormatSwitches MergeFormat = new FieldFormatSwitches(14, "MERGEFORMAT");

        /// <summary>
        /// The value of the field format switch enumeration constant.
        /// </summary>
        private readonly int value;

        /// <summary>
        /// The name of the field format switch enumeration constant.
        /// </summary>
        private readonly string name;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFormatSwitches"/> class.
        /// </summary>
        /// <param name="value">The value of the field format switch enumeration constant.</param>
        /// <param name="name">The name of the field format switch enumeration constant.</param>
        private FieldFormatSwitches(int value, string name)
        {
            this.value = value;
            this.name = FieldFunctionCode.FormatFieldSwitch + name;
        }

        /// <summary>
        /// Returns a string representation of the field format switch enumeration constant.
        /// </summary>
        /// <returns>The string representation.</returns>
        public override string ToString()
        {
            return this.name;
        }
    }
}
