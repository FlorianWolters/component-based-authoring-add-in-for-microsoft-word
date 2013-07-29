//------------------------------------------------------------------------------
// <copyright file="EventHandlerFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using FlorianWolters.Office.Word.Commands;
    using Microsoft.Office.Interop.Word;

    public abstract class EventHandlerFactory : IEventHandlerFactory
    {
        public IEventHandler RegisterEventHandler(
            ApplicationEventHandler applicationEventHandler)
        {
            ICommand command = this.CreateCommand(applicationEventHandler.Application);
            IEventHandler eventHandler = this.CreateEventHandler(command);
            applicationEventHandler.SubscribeEventHandler(eventHandler);

            return eventHandler;
        }

        protected abstract ICommand CreateCommand(Application application);

        protected abstract IEventHandler CreateEventHandler(ICommand command);
    }
}
