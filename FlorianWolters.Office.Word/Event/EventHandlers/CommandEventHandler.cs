//------------------------------------------------------------------------------
// <copyright file="CommandEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.EventHandlers
{
    using System;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;

    public abstract class CommandEventHandler : IEventHandler
    {
        /// <summary>
        /// The <i>Command</i> of this <see cref="CommandEventHandler"/>.
        /// </summary>
        protected readonly ICommand Command;

        /// <summary>
        /// The exception handler of this <see cref="CommandEventHandler"/>.
        /// </summary>
        protected readonly IEventExceptionHandler ExceptionHandler;

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandEventHandler"/>
        /// class with the specified <i>Command</i> and the specified exception
        /// handler.
        /// </summary>
        /// <param name="command">The <i>Command</i>.</param>
        /// <param name="exceptionHandler">The exception handler.</param>
        protected CommandEventHandler(
            ICommand command,
            IEventExceptionHandler exceptionHandler)
        {
            this.Command = command;
            this.ExceptionHandler = exceptionHandler;
        }

        /// <summary>
        /// Tries to execute the <i>Command</i> of this <see
        /// cref="CommandEventHandler"/>.
        /// </summary>
        /// <remarks>
        /// If the <i>Command</i> throws an exception, the exception handler,
        /// configured via the constructor, is used to handle the exception.
        /// </remarks>
        protected void TryExecute()
        {
            try
            {
                this.Command.Execute();
            }
            catch (Exception ex)
            {
                this.ExceptionHandler.HandleException(ex);
            }
        }
    }
}
