//------------------------------------------------------------------------------
// <copyright file="FieldFilePathTranslator.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System.IO;

    /// <summary>
    /// The class <see cref="FieldFilePathTranslator"/> allows to encode and decode a file path representation for a
    /// field in a Microsoft Word document.
    /// </summary>
    public class FieldFilePathTranslator
    {
        /// <summary>
        /// The directory separator used in a field.
        /// </summary>
        public const string DirectorySeparator = @"\\";

        /// <summary>
        /// Encodes the specified normal file path to a field file path.
        /// </summary>
        /// <param name="normalFilePath">The normal file path to encode.</param>
        /// <returns>The field file path.</returns>
        public string Encode(string normalFilePath)
        {
            return normalFilePath.Replace(Path.DirectorySeparatorChar.ToString(), DirectorySeparator);
        }

        /// <summary>
        /// Decodes the specified field file path to a normal file path.
        /// </summary>
        /// <param name="fieldFilePath">The field file path to decode.</param>
        /// <returns>The normal file path.</returns>
        public string Decode(string fieldFilePath)
        {
            return fieldFilePath.Replace(DirectorySeparator, Path.DirectorySeparatorChar.ToString());
        }
    }
}
