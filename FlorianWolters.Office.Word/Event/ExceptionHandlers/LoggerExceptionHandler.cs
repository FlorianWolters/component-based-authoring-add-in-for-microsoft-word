//------------------------------------------------------------------------------
// <copyright file="LoggerExceptionHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.ExceptionHandlers
{
    using System;
    using NLog;

    /// <summary>
    /// The class <see cref="LoggerExceptionHandler"/> uses a <see
    /// cref="Logger"/> to log <see cref="Exception"/> messages.
    /// </summary>
    public class LoggerExceptionHandler : IExceptionHandler
    {
        /// <summary>
        /// The <see cref="Logger"/> to use to log the <see cref="Exception"/>
        /// messages.
        /// </summary>
        private readonly Logger logger;

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="LoggerExceptionHandler"/> class with the specified <see cref="Logger"/>.
        /// </summary>
        /// <param name="logger">The <see cref="Logger"/> to use to log the
        /// <see cref="Exception"/> messages.</param>
        public LoggerExceptionHandler(Logger logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Handles the specified <see cref="Exception"/>.
        /// </summary>
        /// <param name="ex">The <see cref="Exception"/> to handle.</param>
        public void HandleException(Exception ex)
        {
            this.logger.WarnException(ex.Message, ex);
        }
    }
}
