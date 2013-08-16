//------------------------------------------------------------------------------
// <copyright file="IDocumentBeforeCloseEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IDocumentBeforeCloseEventHandler"/> allows to handle the event which occurs immediately
    /// before any open <see cref="Word.Document"/> closes.
    /// </summary>
    public interface IDocumentBeforeCloseEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs immediately before any open <see cref="Word.Document"/> closes.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being closed.</param>
        /// <param name="target">
        /// <c>false</c> when the event occurs. If the event procedure sets this argument to <c>true</c>, the <see
        /// cref="Word.Document"/> doesn't close when the procedure is finished.
        /// </param>
        void OnDocumentBeforeClose(Word.Document document, ref bool target);
    }
}
