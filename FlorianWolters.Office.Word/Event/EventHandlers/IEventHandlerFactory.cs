//------------------------------------------------------------------------------
// <copyright file="IEventHandlerFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;

    public interface IEventHandlerFactory
    {
        IEventHandler RegisterEventHandler(
            IExceptionHandler exceptionHandler,
            ApplicationEventHandler applicationEventHandler);
    }
}
