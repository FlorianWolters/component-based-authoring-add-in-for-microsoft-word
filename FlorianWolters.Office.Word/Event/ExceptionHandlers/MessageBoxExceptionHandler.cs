//------------------------------------------------------------------------------
// <copyright file="MessageBoxExceptionHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event.ExceptionHandlers
{
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="MessageBoxExceptionHandler"/> displays <see
    /// cref="Exception"/> messages in a <see cref="MessageBox"/>.
    /// </summary>
    /// <remarks>
    /// A <see cref="MessageBox"/> blocks the thread it is running in, until it
    /// is closed.
    /// </remarks>
    public class MessageBoxExceptionHandler : IExceptionHandler
    {
        /// <summary>
        /// The owner window of the <see cref="MessageBox"/>.
        /// </summary>
        private readonly IWin32Window owner;

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="MessageBoxExceptionHandler"/> class without a owner
        /// window.
        /// </summary>
        public MessageBoxExceptionHandler() : this(null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="MessageBoxExceptionHandler"/> class with the specified
        /// owner window.
        /// </summary>
        /// <param name="window">The owner window of the <see cref="MessageBox"/>.</param>
        public MessageBoxExceptionHandler(IWin32Window window)
        {
            this.owner = window;
        }

        /// <summary>
        /// Handles the specified <see cref="Exception"/>.
        /// </summary>
        /// <param name="ex">The <see cref="Exception"/> to handle.</param>
        public void HandleException(Exception ex)
        {
            MessageBox.Show(
                this.owner,
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }
}
