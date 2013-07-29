//------------------------------------------------------------------------------
// <copyright file="DocumentExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Extensions
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The static class <see cref="DocumentExtensions"/> contains extension
    /// methods for a Microsoft Word document, represented by an object of the
    /// class <see cref="Word.Document"/>.
    /// </summary>
    public static class DocumentExtensions
    {
        /// <summary>
        /// Determines whether the specified <see cref="Word.Document"/> is
        /// saved.
        /// </summary>
        /// <param name="document">
        /// The <see cref="Word.Document"/> to check.
        /// </param>
        /// <returns>
        /// <c>true</c> if the <see cref="Word.Document"/> is saved,
        /// <c>false</c> otherwise.
        /// </returns>
        public static bool IsSaved(this Word.Document document)
        {
            return string.Empty != document.Path;
        }

        /// <summary>
        /// Updates all Table of Contents (ToC) in the specified <see
        /// cref="Word.Document"/>.
        /// </summary>
        /// <param name="document">
        /// The <see cref="Word.Document"/> to modify.
        /// </param>
        public static void UpdateAllTableOfContents(this Word.Document document)
        {
            if (document.TablesOfContents.Count > 0)
            {
                foreach (Word.TableOfContents table in document.TablesOfContents)
                {
                    table.Update();
                }
            }
        }

        /// <summary>
        /// Updates the Table of Contents (ToC) with the specified index in the
        /// specified <see cref="Word.Document"/>.
        /// </summary>
        /// <param name="document">
        /// The <see cref="Word.Document"/> to modify.
        /// </param>
        /// <param name="index">The index of the ToC to update.</param>
        public static void UpdateTableOfContents(this Word.Document document, int index = 1)
        {
            if (document.TablesOfContents.Count > (index - 1))
            {
                document.TablesOfContents[index].Update();
            }
        }
    }
}
