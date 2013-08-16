//------------------------------------------------------------------------------
// <copyright file="IWindowSelectionChangeEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IWindowSelectionChangeEventHandler"/> allows to handle the event which occurs when the
    /// <see cref="Word.Selection"/> changes in the active <see cref="Word.Document"/> <see cref="Word.Window"/>.
    /// </summary>
    public interface IWindowSelectionChangeEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when the <see cref="Word.Selection"/> changes in the active <see
        /// cref="Word.Document"/> <see cref="Word.Window"/>.
        /// </summary>
        /// <param name="selection">
        /// The text selected. If no text is selected, this parameter is either <c>null</c> or contains the first
        /// character to the right of the insertion point.
        /// </param>
        void OnWindowSelectionChange(Word.Selection selection);
    }
}
