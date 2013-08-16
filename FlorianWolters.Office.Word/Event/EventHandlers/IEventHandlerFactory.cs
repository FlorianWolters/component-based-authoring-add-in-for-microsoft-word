//------------------------------------------------------------------------------
// <copyright file="IEventHandlerFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;

    /// <summary>
    /// Allows to create an <i>Event Handler</i> instance which implements <see cref="IEventHandler"/>.
    /// </summary>
    public interface IEventHandlerFactory
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
        IEventHandler RegisterEventHandler(
            IExceptionHandler exceptionHandler,
            ApplicationEventHandler applicationEventHandler);
    }
}
