//------------------------------------------------------------------------------
// <copyright file="CustomXMLPartDefaultNamespaceException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.CustomXML
{
    using System;

    /// <summary>
    /// The exception <see cref="CustomXMLPartDefaultNamespaceException"/> can be thrown if a custom XML part has a
    /// invalid default namespace.
    /// </summary>
    public class CustomXMLPartDefaultNamespaceException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CustomXMLPartDefaultNamespaceException"/> class with the
        /// specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public CustomXMLPartDefaultNamespaceException(string message)
            : base(message)
        {
        }
    }
}
