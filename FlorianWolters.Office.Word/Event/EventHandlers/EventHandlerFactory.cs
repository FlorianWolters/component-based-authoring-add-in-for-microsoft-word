//------------------------------------------------------------------------------
// <copyright file="EventHandlerFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The abstract class <see cref="EventHandlerFactory"/> allows to create <i>Event Handler</i> instances.
    /// </summary>
    public abstract class EventHandlerFactory : IEventHandlerFactory
    {
        /// <summary>
        /// Creates an <i>Event Handler</i> with the specified exception handler and registers it at the specified <see
        /// cref="ApplicationEventHandler"/>
        /// </summary>
        /// <param name="exceptionHandler">
        /// Used if an <see cref="Exception"/> inside an <i>Event Handler</i> occurs.
        /// </param>
        /// <param name="applicationEventHandler">
        /// Used to register the newly created <i>Event Handler</i> at the Microsoft Word application.
        /// </param>
        /// <returns>The newly created <i>Event Handler</i> instance.</returns>
        public IEventHandler RegisterEventHandler(
            IExceptionHandler exceptionHandler,
            ApplicationEventHandler applicationEventHandler)
        {
            ICommand command = this.CreateCommand(applicationEventHandler.Application);
            IEventHandler eventHandler = this.CreateEventHandler(command, exceptionHandler);
            applicationEventHandler.SubscribeEventHandler(eventHandler);

            return eventHandler;
        }

        /// <summary>
        /// Creates the <i>Command</i> to inject into the <i>Event Handler</i>.
        /// </summary>
        /// <param name="application">The Microsoft Word application used by the <i>Command</i>.</param>
        /// <returns>The newly created <i>Command</i> instance.</returns>
        protected abstract ICommand CreateCommand(Word.Application application);

        /// <summary>
        /// Creates the <i>Event Handler</i> to return by the <i>Factory Method</i>.
        /// </summary>
        /// <param name="command">The <i>Command</i> to inject into the <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">
        /// The exception handler to use if an <see cref="Exception"/> inside an <i>Event Handler</i> occurs.
        /// </param>
        /// <returns>The newly created <i>Event Handler</i> instance.</returns>
        protected abstract IEventHandler CreateEventHandler(ICommand command, IExceptionHandler exceptionHandler);
    }
}
