//------------------------------------------------------------------------------
// <copyright file="IDocumentBeforePrintEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IDocumentBeforeCloseEventHandler"/> allows to handle the event which occurs before any
    /// open <see cref="Word.Document"/> is printed.
    /// </summary>
    public interface IDocumentBeforePrintEventHandler : IEventHandler
    {
        /// <summary>
        /// Handle the event which occurs before any open <see cref="Word.Document"/> is printed.
        /// </summary>
        /// <param name="document">The see <see cref="Word.Document"/> that's being printed.</param>
        /// <param name="cancel">
        /// <c>false</c> when the event occurs. If the event procedure sets this argument to <c>true</c>, the <see
        /// cref="Word.Document"/> isn't printed when the procedure is finished.
        /// </param>
        void OnDocumentBeforePrint(Word.Document document, ref bool cancel);
    }
}
