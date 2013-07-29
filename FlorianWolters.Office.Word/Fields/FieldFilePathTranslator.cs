//------------------------------------------------------------------------------
// <copyright file="FieldFilePathTranslator.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System.IO;

    public class FieldFilePathTranslator
    {
        public const string DirectorySeparator = @"\\";

        public string Encode(string filePath)
        {
            return filePath.Replace(
                Path.DirectorySeparatorChar.ToString(),
                DirectorySeparator);
        }

        public string Decode(string fieldFilePath)
        {
            return fieldFilePath.Replace(
                DirectorySeparator,
                Path.DirectorySeparatorChar.ToString());
        }
    }
}
