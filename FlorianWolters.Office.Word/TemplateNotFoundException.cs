//------------------------------------------------------------------------------
// <copyright file="TemplateNotFoundException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word
{
    using System;
    using System.IO;

    /// <summary>
    /// The exception <see cref="TemplateNotFoundException"/> can be thrown when a Microsoft Word template file does not
    /// exist.
    /// </summary>
    [Serializable]
    public class TemplateNotFoundException : FileNotFoundException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TemplateNotFoundException"/> class with the specified error
        /// message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public TemplateNotFoundException(string message)
            : base(message)
        {
        }
    }
}
