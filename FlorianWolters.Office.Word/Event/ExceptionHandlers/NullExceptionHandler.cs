//------------------------------------------------------------------------------
// <copyright file="NullEventExceptionHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.ExceptionHandlers
{
    using System;

    public class NullEventExceptionHandler : IEventExceptionHandler
    {
        public void HandleException(Exception ex)
        {
            // NOOP
        }
    }
}
