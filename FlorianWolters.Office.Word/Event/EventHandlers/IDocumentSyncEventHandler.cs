//------------------------------------------------------------------------------
// <copyright file="IDocumentSyncEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    public interface IDocumentSyncEventHandler : IEventHandler
    {
        void OnDocumentSync();
    }
}
