//------------------------------------------------------------------------------
// <copyright file="IWindowSizeEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IWindowSizeEventHandler"/> allows to handle the event which occurs when the <see
    /// cref="Word.Application"/> <see cref="Word.Window"/> is resized or moved.
    /// </summary>
    public interface IWindowSizeEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when the <see cref="Word.Application"/> <see cref="Word.Window"/> is resized
        /// or moved.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> in the <see cref="Word.Window"/> being sized.</param>
        /// <param name="window">The <see cref="Word.Window"/> being sized.</param>
        void OnWindowSize(Word.Document document, Word.Window window);
    }
}
