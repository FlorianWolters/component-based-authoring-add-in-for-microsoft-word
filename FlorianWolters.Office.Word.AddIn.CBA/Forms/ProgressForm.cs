//------------------------------------------------------------------------------
// <copyright file="ProgressForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="ProgressForm"/> implements a simple dialog with a
    /// label and a marque progress bar.
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
        /// Raises the <see cref="FormClosing"/> event.
        /// <para>
        /// Prevents closing of this <see cref="ProgressForm"/> for the user.
        /// </para>
        /// </summary>
        /// <param name="e">A <see cref="FormClosingEventArgs"/> that contains
        /// the event data.</param>
        /// <remarks>
        /// The code has been taken from <a
        /// href="http://stackoverflow.com/questions/14943/how-to-disable-alt-f4-closing-form">this</a>
        /// Stack Overflow question.
        /// </remarks>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason.Equals(CloseReason.UserClosing))
            {
                e.Cancel = true;
            }

            base.OnFormClosing(e);
        }
    }
}
