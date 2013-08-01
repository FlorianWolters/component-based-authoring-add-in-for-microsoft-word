//------------------------------------------------------------------------------
// <copyright file="NullExceptionHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.ExceptionHandlers
{
    using System;

    /// <summary>
    /// The class <see cref="NullExceptionHandler"/> does nothing to handle a
    /// <see cref="Exception"/>.
    /// </summary>
    /// <remarks>This class can be used for automatic testing.</remarks>
    public class NullExceptionHandler : IExceptionHandler
    {
        /// <summary>
        /// Handles the specified <see cref="Exception"/>.
        /// </summary>
        /// <param name="ex">The <see cref="Exception"/> to handle.</param>
        public void HandleException(Exception ex)
        {
            // NOOP
        }
    }
}
