//------------------------------------------------------------------------------
// <copyright file="INewDocumentEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="INewDocumentEventHandler"/> allows to handle the event which occurs when a new <see
    /// cref="Word.Document"/> is created.
    /// </summary>
    public interface INewDocumentEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when a new <see cref="Word.Document"/> is created.
        /// </summary>
        /// <param name="document">The new <see cref="Word.Document"/>.</param>
        void OnNewDocument(Word.Document document);
    }
}
