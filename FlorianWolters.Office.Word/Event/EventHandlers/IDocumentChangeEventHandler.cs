//------------------------------------------------------------------------------
// <copyright file="IDocumentChangeEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    /// <summary>
    /// The interface <see cref="IDocumentChangeEventHandler"/> allows to handle the event which occurs when a new <see
    /// cref="Microsoft.Office.Interop.Word.Document"/> is created, when an existing <see
    /// cref="Microsoft.Office.Interop.Word.Document"/> is opened, or when another <see
    /// cref="Microsoft.Office.Interop.Word.Document"/> is made the active <see
    /// cref="Microsoft.Office.Interop.Word.Document"/>.
    /// </summary>
    public interface IDocumentChangeEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when a new <see cref="Microsoft.Office.Interop.Word.Document"/> is created,
        /// when an existing <see cref="Microsoft.Office.Interop.Word.Document"/> is opened, or when another <see
        /// cref="Microsoft.Office.Interop.Word.Document"/> is made the active <see
        /// cref="Microsoft.Office.Interop.Word.Document"/>. 
        /// </summary>
        void OnDocumentChange();
    }
}
