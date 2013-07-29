//------------------------------------------------------------------------------
// <copyright file="ApplicationExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Extensions
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The static class <see cref="ApplicationExtensions"/> contains extension
    /// methods for a Microsoft Word application, represented by an object of
    /// the class <see cref="Word.Application"/>.
    /// </summary>
    public static class ApplicationExtensions
    {
        /// <summary>
        /// Determines whether at least one document is opened in the specified
        /// <see cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">
        /// The <see cref="Word.Application"/> to check.
        /// </param>
        /// <returns>
        /// <c>true</c> if at least one document is opened, <c>false</c>
        /// otherwise.
        /// </returns>
        public static bool HasOpenDocuments(this Word.Application application)
        {
            return application.Documents.Count > 0;
        }
    }
}
