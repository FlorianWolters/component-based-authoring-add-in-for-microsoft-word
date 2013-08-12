//------------------------------------------------------------------------------
// <copyright file="MessagesForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System.Windows.Forms;

    public partial class MessagesForm : Form
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
        /// Handles the <see cref="FormClosing"/> event of this <see cref="MessagesForm"/>.
        /// <para>Prevents the closing of the form by the user.</para>
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="FormClosingEventArgs"/> that contains the event data.</param>
        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = (e.CloseReason == CloseReason.UserClosing);
        }
    }
}
