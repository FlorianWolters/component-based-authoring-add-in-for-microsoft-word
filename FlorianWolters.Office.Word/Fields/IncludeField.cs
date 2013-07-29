//------------------------------------------------------------------------------
// <copyright file="IncludeField.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System;
    using System.Text.RegularExpressions;
    using FlorianWolters.Office.Word.Extensions;
    using Word = Microsoft.Office.Interop.Word;

    public class IncludeField
    {
        private readonly Word.Field field;

        public IncludeField(Word.Field field)
        {
            if (null == field)
            {
                throw new ArgumentNullException("field");
            }

            if (!field.CanUpdate())
            {
                throw new ArgumentException("field");
            }

            this.field = field;
        }

        public string FilePath
        {
            get
            {
                const string Pattern = "INCLUDE(?:TEXT|PICTURE)?.+\\s+\"(.+)\"";

                Word.Range fieldRange = this.field.Code;
                fieldRange.TextRetrievalMode.IncludeFieldCodes = false;

                Match match = Regex.Match(
                    fieldRange.Text,
                    Pattern,
                    RegexOptions.IgnoreCase);

                if (!match.Success)
                {
                    throw new FormatException(
                        "Unable to retrieve the file path from the INCLUDE field.");
                }

                // TODO Fix violation of IoC.
                return new FieldFilePathTranslator().Decode(match.Groups[1].Value);
            }
        }
    }
}
