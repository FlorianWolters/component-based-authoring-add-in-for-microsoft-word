//------------------------------------------------------------------------------
// <copyright file="IWindowActivateEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IWindowActivateEventHandler"/> allows to handle the event which occurs when any <see
    /// cref="Word.Document"/> <see cref="Word.Window"/> is activated.
    /// </summary>
    public interface IWindowActivateEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when any <see cref="Word.Document"/> <see cref="Word.Window"/> is activated.
        /// </summary>
        /// <param name="document">
        /// The <see cref="Word.Document"/> displayed in the activated <see cref="Word.Window"/>.
        /// </param>
        /// <param name="window">The <see cref="Word.Window"/> that's being activated.</param>
        void OnWindowActivate(Word.Document document, Word.Window window);
    }
}
