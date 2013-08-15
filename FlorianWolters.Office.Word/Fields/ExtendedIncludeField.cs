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
    /// This class can get and set data from a Microsoft Word field which has two nested fields, one <c>DocProperty</c>
    /// field (which holds the absolute base directory path for the included file) and one empty field (which holds the
    /// date and time of the last modification of the included file).
    /// </para>
    /// </summary>
    public class ExtendedIncludeField
    {
        /// <summary>
        /// A standard date and time format string.
        /// <para>
        /// The universal sortable date/time pattern is used.
        /// </para>
        /// </summary>
        private const string DateTimeFormat = "u";

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedIncludeField"/> class.
        /// </summary>
        /// <param name="field">A <c>IncludeText</c>, <c>IncludePicture</c> or <c>Include</c> <see cref="Word.Field"/>.</param>
        public ExtendedIncludeField(Word.Field field)
        {
            if (null == field)
            {
                throw new ArgumentNullException("field");
            }

            if (!field.IsTypeInclude())
            {
                // The field has the wrong type.
                throw new ArgumentException("field");
            }

            Word.Fields nestedFields = field.Code.Fields;

            if (2 != nestedFields.Count
                || Word.WdFieldType.wdFieldDocProperty != nestedFields[1].Type
                || Word.WdFieldType.wdFieldEmpty != nestedFields[2].Type)
            {
                // The field doesn't have two nested fields.
                // The first nested field must be of type DocProperty and the second nested field must be of type empty.
                throw new FormatException("field");
            } 

            this.IncludeField = field;
            this.DocPropertyField = nestedFields[1];
            this.EmptyField = nestedFields[2];
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
        /// Gets the nested empty <see cref="Word.Field"/>.
        /// </summary>
        public Word.Field EmptyField { get; private set; }

        /// <summary>
        /// Gets the (absolute) file path of this <see cref="ExtendedIncludeField"/>.
        /// </summary>
        public string FilePath
        {
            get
            {
                const string Pattern = "INCLUDE(?:TEXT|PICTURE)?.+\\s+\"(.+)\"";

                Word.Range fieldRange = this.IncludeField.Code;
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

        /// <summary>
        /// Gets the date and time of the last modification of this <see cref="ExtendedIncludeField"/>.
        /// </summary>
        public string LastModified
        {
            get
            {
                return this.EmptyField.Code.Text.Trim();
            }

            private set
            {
                this.EmptyField.Code.Text = value;
            }
        }

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
        /// Updates the date and time of the empty <see cref="Word.Field"/> to the date and time of the last
        /// modification of the included source file.
        /// </summary>
        public void SynchronizeLastModified()
        {
            this.LastModified = LastModifiedForFile(this.FilePath);
        }

        /// <summary>
        /// Determines whether this <see cref="ExtendedIncludeField"/> is in sync with the referenced source file.
        /// </summary>
        /// <returns><c>true</c> if in sync; <c>false</c> if out-of-sync.</returns>
        public bool IsInSync()
        {
            return ExtendedIncludeField.LastModifiedForFile(this.FilePath) == this.LastModified;
        }

        /// <summary>
        /// Determines whether this <see cref="ExtendedIncludeField"/> is out-of-sync with the referenced source file.
        /// </summary>
        /// <returns><c>true</c> if out-of-sync; <c>false</c> if in sync.</returns>
        public bool IsNotInSync()
        {
            return !this.IsInSync();
        }

        /// <summary>
        /// Returns the date and time, in coordinated universal time (UTC), that the specified file or directory was
        /// last written to.
        /// </summary>
        /// <param name="filePath">The file or directory for which to obtain write date and time information.</param>
        /// <returns>A standard date and time format string.</returns>
        private static string LastModifiedForFile(string filePath)
        {
            return File.GetLastWriteTimeUtc(filePath).ToString(DateTimeFormat);
        }
    }
}
