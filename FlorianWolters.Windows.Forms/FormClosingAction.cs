//------------------------------------------------------------------------------
// <copyright file="FormClosingAction.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------
namespace FlorianWolters.Windows.Forms
{
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="FormClosingAction"/> allows interaction with a <see
    /// cref="FormClosingEventArgs"/> object.
    /// </summary>
    public class FormClosingAction
    {
        /// <summary>
        /// A <see cref="FormClosingEventArgs"/> that contains the event data.
        /// </summary>
        private readonly FormClosingEventArgs e;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormClosingAction"/>
        /// class with the specified <see cref="FormClosingEventArgs"/>.
        /// </summary>
        /// <param name="e">A <see cref="FormClosingEventArgs"/> that contains
        /// the event data.</param>
        public FormClosingAction(FormClosingEventArgs e)
        {
            this.e = e;
        }

        /// <summary>
        /// Prevents closing of the form.
        /// </summary>
        /// <remarks>
        /// The code has been taken from <a
        /// href="http://stackoverflow.com/questions/14943/how-to-disable-alt-f4-closing-form">this</a>
        /// Stack Overflow question and modified by the author of this class.
        /// </remarks>
        public void Cancel()
        {
            if (this.e.CloseReason.Equals(CloseReason.UserClosing))
            {
                this.e.Cancel = true;
            }
        }
    }
}
