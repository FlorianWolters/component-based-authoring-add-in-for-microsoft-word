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
    /// The interface <see cref="IDocumentBeforeSaveEventHandler"/> allows to handle the event which occurs before any
    /// <see cref="Word.Document"/> is saved.
    /// </summary>
    public interface IDocumentBeforeSaveEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs before any <see cref="Word.Document"/> is saved.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being saved.</param>
        /// <param name="saveAsUI">
        /// <c>true</c> if the <b>Save As</b> dialog box is displayed, whether to save a new <see
        /// cref="Word.Document"/>, in response to the <b>Save</b> command; or in response to the <b>Save As</b>
        /// command; or in response to the <b>SaveAs</b> or <b>SaveAs2</b> method.
        /// </param>
        /// <param name="cancel">
        /// <c>false</c> when the event occurs. If the event procedure sets this argument to <c>true</c>, the <see
        /// cref="Word.Document"/> is not saved when the procedure is finished.
        /// </param>
        void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel);
    }
}
