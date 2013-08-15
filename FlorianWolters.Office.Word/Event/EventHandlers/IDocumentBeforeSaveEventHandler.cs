//------------------------------------------------------------------------------
// <copyright file="IDocumentBeforeSaveEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IDocumentBeforeSaveEventHandler"/> allows to handle the event which occurs before a
    /// <see cref="Word.Document"/> is saved.
    /// </summary>
    public interface IDocumentBeforeSaveEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs before a <see cref="Word.Document"/> is saved.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being saved.</param>
        /// <param name="saveAsUI"><c>true</c> if called via <i>Save As</i>; <c>false</c> if called via <i>Save</i>.</param>
        /// <param name="cancel"><c>true</c> to prevent the <see cref="Word.Document"/> from being saved; <c>false</c> otherwise.</param>
        void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel);
    }
}
