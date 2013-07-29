//------------------------------------------------------------------------------
// <copyright file="CustomXMLPartDefaultNamespaceException.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.CustomXML
{
    using System;

    public class CustomXMLPartDefaultNamespaceException : Exception
    {
        public CustomXMLPartDefaultNamespaceException(string message)
            : base(message)
        {
        }
    }
}
