//------------------------------------------------------------------------------
// <copyright file="IStartupEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    public interface IStartupEventHandler : IEventHandler
    {
        void OnStartup();
    }
}
