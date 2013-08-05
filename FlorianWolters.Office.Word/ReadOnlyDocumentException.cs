//------------------------------------------------------------------------------
// <copyright file="ReadOnlyDocumentException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word
{
    using System;

    /// <summary>
    /// The exception <see cref="ReadOnlyDocumentException"/> can be thrown if a
    /// Microsoft Word document is read-only.
    /// </summary>
    public class ReadOnlyDocumentException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="ReadOnlyDocumentException"/> class with the specified error
        /// message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ReadOnlyDocumentException(string message) : base(message)
        {
        }
    }
}
