//------------------------------------------------------------------------------
// <copyright file="FieldCreationException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.DocumentProperties
{
    using System;

    public class FieldCreationException : Exception
    {
        public FieldCreationException(string message)
            : base(message)
        {
        }
    }
}
