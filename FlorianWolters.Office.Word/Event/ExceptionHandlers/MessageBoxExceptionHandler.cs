//------------------------------------------------------------------------------
// <copyright file="MessageBoxEventExceptionHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.ExceptionHandlers
{
    using System;
    using System.Windows.Forms;

    public class MessageBoxEventExceptionHandler : IEventExceptionHandler
    {
        public void HandleException(Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
