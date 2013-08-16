//------------------------------------------------------------------------------
// <copyright file="IXMLValidationErrorEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IXMLValidationErrorEventHandler"/> allows to handle the event which occurs when there
    /// is a validation error in the <see cref="Word.Document"/>.
    /// </summary>
    public interface IXMLValidationErrorEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when there is a validation error in the <see cref="Word.Document"/>.
        /// </summary>
        /// <param name="xmlNode">The XML element that is invalid.</param>
        void OnXMLValidationError(Word.XMLNode xmlNode);
    }
}
