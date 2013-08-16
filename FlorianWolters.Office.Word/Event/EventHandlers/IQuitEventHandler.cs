//------------------------------------------------------------------------------
// <copyright file="IQuitEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    /// <summary>
    /// The interface <see cref="IQuitEventHandler"/> allows to handle the event which occurs when the user exits
    /// Microsoft Word.
    /// </summary>
    public interface IQuitEventHandler : IEventHandler
    {
        /// <summary>
        /// Handles the event which occurs when the user exits Microsoft Word.
        /// </summary>
        void OnQuit();
    }
}
