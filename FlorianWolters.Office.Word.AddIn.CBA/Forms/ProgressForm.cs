//------------------------------------------------------------------------------
// <copyright file="ProgressForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System.Windows.Forms;
    using FlorianWolters.Windows.Forms;

    /// <summary>
    /// The class <see cref="ProgressForm"/> implements a simple progress form with a label and a marque progress bar.
    /// </summary>
    internal partial class ProgressForm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ProgressForm"/> class.
        /// </summary>
        public ProgressForm()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Changes the text for the label of this <see cref="ProgressForm"/>.
        /// </summary>
        /// <param name="text">The new text for the label.</param>
        public void ChangeLabelText(string text)
        {
            this.label.Text = text;
        }

        /// <summary>
        /// Closes this <see cref="ProgressForm"/>.
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
        /// Handles the <see cref="Form.FormClosing"/> event of this <see cref="ProgressForm"/>.
        /// <para>Prevents the closing of the by the user.</para>
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="FormClosingEventArgs"/> that contains the event data.</param>
        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = e.CloseReason == CloseReason.UserClosing;
        }
    }
}
