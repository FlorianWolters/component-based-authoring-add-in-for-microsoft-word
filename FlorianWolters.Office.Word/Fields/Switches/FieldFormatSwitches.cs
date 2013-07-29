//------------------------------------------------------------------------------
// <copyright file="FieldFormatSwitches.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.Switches
{
    using FlorianWolters.Office.Word.Fields;

    public sealed class FieldFormatSwitches
    {
        public static readonly FieldFormatSwitches Caps = new FieldFormatSwitches(1, "CAPS");
        public static readonly FieldFormatSwitches FirstCap = new FieldFormatSwitches(2, "FIRSTCAP");
        public static readonly FieldFormatSwitches Upper = new FieldFormatSwitches(3, "UPPER");
        public static readonly FieldFormatSwitches Lower = new FieldFormatSwitches(4, "LOWER");
        public static readonly FieldFormatSwitches Alphabetic = new FieldFormatSwitches(5, "ALPHABETIC");
        public static readonly FieldFormatSwitches Arabic = new FieldFormatSwitches(6, "ARABIC");
        public static readonly FieldFormatSwitches CardText = new FieldFormatSwitches(7, "CARDTEXT");
        public static readonly FieldFormatSwitches DollarText = new FieldFormatSwitches(8, "DOLLARTEXT");
        public static readonly FieldFormatSwitches Hex = new FieldFormatSwitches(9, "HEX");
        public static readonly FieldFormatSwitches OrdText = new FieldFormatSwitches(10, "ORDTEXT");
        public static readonly FieldFormatSwitches Ordinal = new FieldFormatSwitches(11, "ORDINAL");
        public static readonly FieldFormatSwitches Roman = new FieldFormatSwitches(12, "ROMAN");
        public static readonly FieldFormatSwitches CharFormat = new FieldFormatSwitches(13, "CHARFORMAT");
        public static readonly FieldFormatSwitches MergeFormat = new FieldFormatSwitches(14, "MERGEFORMAT");

        private readonly int value;
        private readonly string name;

        private FieldFormatSwitches(int value, string name)
        {
            this.value = value;
            this.name = FieldFunctionCode.FormatFieldSwitch + name;
        }

        public override string ToString()
        {
            return this.name;
        }
    }
}
