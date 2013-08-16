//------------------------------------------------------------------------------
// <copyright file="IWindowDeactivateEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IWindowDeactivateEventHandler"/> allows to handle the event which occurs when any <see
    /// cref="Word.Document"/> <see cref="Word.Window"/> is deactivated.
    /// </summary>
    public interface IWindowDeactivateEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when any <see cref="Word.Document"/> <see cref="Word.Window"/> is
        /// deactivated.
        /// </summary>
        /// <param name="document">
        /// The <see cref="Word.Document"/> displayed in the deactivated <see cref="Word.Window"/>.
        /// </param>
        /// <param name="window">The deactivated <see cref="Word.Window"/>.</param>
        void OnWindowDeactivate(Word.Document document, Word.Window window);
    }
}
