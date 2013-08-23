//------------------------------------------------------------------------------
// <copyright file="FieldCreationException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.DocumentProperties
{
    using System;

    /// <summary>
    /// The exception <see cref="FieldCreationException"/> can be thrown if a field cannot be created in a Microsoft
    /// Word document.
    /// </summary>
    [Serializable]
    public class FieldCreationException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FieldCreationException"/> class with the specified error
        /// message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public FieldCreationException(string message)
            : base(message)
        {
        }
    }
}
