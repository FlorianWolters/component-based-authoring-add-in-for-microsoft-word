//------------------------------------------------------------------------------
// <copyright file="MessagesForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="MessagesForm"/> implements a simple dialog, which displays log messages (e.g. for events)
    /// of the application.
    /// </summary>
    /// <remarks>
    /// This class is compatible with <i>NLog</i> and can be specified in the <c>formName</c> argument for a
    /// <c>RichTextBox</c> target.
    /// </remarks>
    internal partial class MessagesForm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MessagesForm" /> class.
        /// </summary>
        public MessagesForm()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Closes this <see cref="MessagesForm"/>.
        /// <para>
        /// This method removes the <c>FormClosing</c> event handler <c>OnFormClosing</c> before closing the form.
        /// </para>
        /// </summary>
        public new void Close()
        {
            this.FormClosing -= this.OnFormClosing;
            base.Close();
        }

        /// <summary>
        /// Handles the <see cref="Form.FormClosing"/> event of this <see cref="MessagesForm"/>.
        /// <para>Prevents the closing of the form by the user.</para>
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="FormClosingEventArgs"/> that contains the event data.</param>
        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = e.CloseReason == CloseReason.UserClosing;
        }
    }
}
