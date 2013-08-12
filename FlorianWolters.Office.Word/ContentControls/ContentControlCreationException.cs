//------------------------------------------------------------------------------
// <copyright file="ContentControlCreationException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.ContentControls
{
    using System;

    /// <summary>
    /// The exception <see cref="ContentControlMappingException"/> can be thrown if the creation of a content control
    /// fails.
    /// </summary>
    public class ContentControlCreationException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ContentControlCreationException"/> class with the specified
        /// error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public ContentControlCreationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentControlCreationException"/> class with a specified error
        /// message and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception. </param>
        /// <param name="innerException">
        /// The exception that is the cause of the current exception, or a <c>null</c> reference if no inner exception
        /// is specified.
        /// </param>
        public ContentControlCreationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
