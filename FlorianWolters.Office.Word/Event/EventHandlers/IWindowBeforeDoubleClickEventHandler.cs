//------------------------------------------------------------------------------
// <copyright file="IWindowBeforeDoubleClickEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IWindowBeforeDoubleClickEventHandler"/> allows to handle the event which occurs when 
    /// the editing area of a <see cref="Word.Document"/> <see cref="Word.Window"/> is double-clicked, before the
    /// default double-click action.
    /// </summary>
    public interface IWindowBeforeDoubleClickEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when the editing area of a <see cref="Word.Document"/> <see
        /// cref="Word.Window"/> is double-clicked, before the default double-click action.
        /// </summary>
        /// <param name="selection">The current <see cref="Word.Selection"/>.</param>
        /// <param name="cancel">
        /// <c>false</c> when the event occurs. If the event procedure sets this argument to <c>true</c>, the default
        /// double-click action does not occur when the procedure is finished.
        /// </param>
        void OnWindowBeforeDoubleClick(Word.Selection selection, ref bool cancel);
    }
}
