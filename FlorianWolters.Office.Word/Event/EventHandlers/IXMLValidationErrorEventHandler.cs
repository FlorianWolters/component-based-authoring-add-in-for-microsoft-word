﻿//------------------------------------------------------------------------------
// <copyright file="IXMLValidationErrorEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    public interface IXMLValidationErrorEventHandler
    {
        void OnXMLValidationError(Word.XMLNode xmlNode);
    }
}
