//------------------------------------------------------------------------------
// <copyright file="UnknownCustomDocumentPropertyException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.DocumentProperties
{
    using System;

    /// <summary>
    /// The exception <see cref="UnknownCustomDocumentPropertyException"/> can
    /// be thrown when an application tries to access an undefined custom
    /// document property of a Microsoft Word document file.
    /// </summary>
    public class UnknownCustomDocumentPropertyException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="UnknownCustomDocumentPropertyException"/> class with the
        /// specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public UnknownCustomDocumentPropertyException(string message)
            : base(message)
        {
        }
    }
}
