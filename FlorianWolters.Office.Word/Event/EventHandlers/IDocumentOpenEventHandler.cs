//------------------------------------------------------------------------------
// <copyright file="IDocumentOpenEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IDocumentOpenEventHandler"/> allows to handle the event which occurs when a <see
    /// cref="Word.Document"/> is opened.
    /// </summary>
    public interface IDocumentOpenEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when a <see cref="Word.Document"/> is opened.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being opened.</param>
        void OnDocumentOpen(Word.Document document);
    }
}
