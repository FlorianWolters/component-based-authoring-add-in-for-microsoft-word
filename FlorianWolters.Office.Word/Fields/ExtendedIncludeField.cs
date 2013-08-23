//------------------------------------------------------------------------------
// <copyright file="ExtendedIncludeField.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System;
    using System.IO;
    using System.Text.RegularExpressions;
    using FlorianWolters.Office.Word.Extensions;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="ExtendedIncludeField"/> extends the functionality of a Microsoft Word <c>IncludeText</c>,
    /// <c>IncludePicture</c> and <c>Include</c> field.
    /// <para>
    /// This class can get and set data from a Microsoft Word field which has one nested fields, a <c>DocProperty</c>
    /// field (which holds the absolute base directory path for the included file).
    /// </para>
    /// </summary>
    public class ExtendedIncludeField
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedIncludeField"/> class.
        /// </summary>
        /// <param name="field">
        /// A <c>IncludeText</c>, <c>IncludePicture</c> or <c>Include</c> <see cref="Word.Field"/>.
        /// </param>
        /// <exception cref="ArgumentNullException">If <c>field</c> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">If the <see cref="Word.Field"/> has the wrong type.</exception>
        /// <exception cref="FormatException">If the <see cref="Word.Field"/> has an invalid format.</exception>
        public ExtendedIncludeField(Word.Field field)
        {
            this.FilePath = this.ParseFilePath(field);

            Word.Fields nestedFields = field.Code.Fields;

            if (1 != nestedFields.Count || Word.WdFieldType.wdFieldDocProperty != nestedFields[1].Type)
            {
                // The field doesn't have a nested DocProperty field.
                throw new FormatException("field");
            } 

            this.IncludeField = field;
            this.DocPropertyField = nestedFields[1];
        }

        /// <summary>
        /// Gets the <c>IncludeText</c>, <c>IncludePicture</c> or <c>Include</c> <see cref="Word.Field"/>.
        /// </summary>
        public Word.Field IncludeField { get; private set; }

        /// <summary>
        /// Gets the nested <c>DocProperty</c> <see cref="Word.Field"/>.
        /// </summary>
        public Word.Field DocPropertyField { get; private set; }

        /// <summary>
        /// Gets the (absolute) file path of this <see cref="ExtendedIncludeField"/>.
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// Tries to create a <see cref="ExtendedIncludeField"/> from the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">A <c>IncludeText</c>, <c>IncludePicture</c> or <c>Include</c> <see cref="Word.Field"/>.</param>
        /// <param name="extendedIncludeField">The <see cref="ExtendedIncludeField"/> to create.</param>
        /// <returns><c>true</c> on success; <c>false</c> on failure.</returns>
        public static bool TryCreateExtendedIncludeField(Word.Field field, out ExtendedIncludeField extendedIncludeField)
        {
            extendedIncludeField = null;
            bool result = true;

            try
            {
                extendedIncludeField = new ExtendedIncludeField(field);
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Returns the file path from the specified <i>INCLUDE[...]</i> <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to parse.</param>
        /// <returns>The file path.</returns>
        /// <exception cref="ArgumentNullException">If <c>field</c> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException">If the <see cref="Word.Field"/> has the wrong type.</exception>
        /// <exception cref="FormatException">If the <see cref="Word.Field"/> has an invalid format.</exception>
        private string ParseFilePath(Word.Field field)
        {
            if (null == field)
            {
                throw new ArgumentNullException("field");
            }

            if (!field.IsTypeInclude())
            {
                throw new ArgumentException("field");
            }

            const string Pattern = "INCLUDE(?:TEXT|PICTURE)?.+\\s+\"(.+)\"";

            Word.Range fieldRange = field.Code;
            fieldRange.TextRetrievalMode.IncludeFieldCodes = false;

            Match match = Regex.Match(
                fieldRange.Text,
                Pattern,
                RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                throw new FormatException("Unable to retrieve the file path from the INCLUDE field.");
            }

            // TODO Fix violation of IoC.
            return new FieldFilePathTranslator().Decode(match.Groups[1].Value);
        }
    }
}
