//------------------------------------------------------------------------------
// <copyright file="ReadOnlyDocumentException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word
{
    using System;

    public class ReadOnlyDocumentException : Exception
    {
        public ReadOnlyDocumentException(string message) : base(message)
        {
        }
    }
}
