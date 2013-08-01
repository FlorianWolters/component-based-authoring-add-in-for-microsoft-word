//------------------------------------------------------------------------------
// <copyright file="IExceptionHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.ExceptionHandlers
{
    using System;

    /// <summary>
    /// The interface <see cref="IExceptionHandler"/> allows an implementing
    /// class to handle a <see cref="Exception"/>.
    /// </summary>
    public interface IExceptionHandler
    {
        /// <summary>
        /// Handles the specified <see cref="Exception"/>.
        /// </summary>
        /// <param name="ex">The <see cref="Exception"/> to handle.</param>
        void HandleException(Exception ex);
    }
}
