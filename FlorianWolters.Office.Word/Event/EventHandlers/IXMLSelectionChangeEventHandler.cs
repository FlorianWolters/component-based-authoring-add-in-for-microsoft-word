//------------------------------------------------------------------------------
// <copyright file="IXMLSelectionChangeEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IXMLSelectionChangeEventHandler"/> allows to handle the event which occurs when the
    /// parent XML node of the current <see cref="Word.Selection"/> changes.
    /// </summary>
    public interface IXMLSelectionChangeEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when the parent XML node of the current <see cref="Word.Selection"/> changes.
        /// </summary>
        /// <param name="selection">
        /// The text selected. If no text is selected, this parameter is either <c>null</c> or contains the first
        /// character to the right of the insertion point.
        /// </param>
        /// <param name="oldXMLNode">The XML node from which the insertion point is moving.</param>
        /// <param name="newXMLNode">The XML node to which the insertion point is moving.</param>
        /// <param name="reason">Can be any of the <see cref="Word.wdXMLSelectionChange"/> constants.</param>
        void OnXMLSelectionChange(
            Word.Selection selection,
            Word.XMLNode oldXMLNode,
            Word.XMLNode newXMLNode,
            ref int reason);
    }
}
