//------------------------------------------------------------------------------
// <copyright file="Dialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Word = Microsoft.Office.Interop.Word;

    public abstract class Dialog
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Dialog"/> class with
        /// the specified <see cref="Word.Application"/> and the specified <see
        /// cref="Word.WdWordDialog"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        /// <param name="dialogType">
        /// The type of the built-in Microsoft Word dialog box.
        /// </param>
        protected Dialog(
            Word.Application application,
            Word.WdWordDialog dialogType)
        {
            this.Application = application;
            this.WordDialog = application.Dialogs[dialogType];
        }

        /// <summary>
        /// Gets the Microsoft Word Application to interact with.
        /// </summary>
        protected Word.Application Application { get; private set; }

        /// <summary>
        /// Gets the Microsoft Word dialog box (represented via a constant of
        /// the <see cref="Word.WdWordDialog"/> enumeration) to handle.
        /// </summary>
        protected dynamic WordDialog { get; private set; }

        /// <summary>
        /// Displays and carries out actions initiated in the specified built-in
        /// Microsoft Word dialog box represented by this <see cref="Dialog"/>.
        /// </summary>
        /// <returns>
        /// An identifier of the enumeration <see cref="DialogResults"/>,
        /// indicating the return value of this <see cref="Dialog"/>.
        /// </returns>
        public DialogResults Show()
        {
            return this.HandleResult(this.Display());
        }

        /// <summary>
        /// Handles the result of this <see cref="Dialog"/>.
        /// <para>
        /// By default, the current settings of the Microsoft Word dialog box
        /// are applied. This method can be overwritten to change that behavior.
        /// </para>
        /// </summary>
        /// <param name="result">
        /// An identifier of the enumeration <see cref="DialogResults"/>,
        /// indicating the return value of the built-in Microsoft Word dialog
        /// box.
        /// </param>
        /// <returns>
        /// An identifier of the enumeration <see cref="DialogResults"/>,
        /// indicating the return value of this <see cref="Dialog"/>.
        /// </returns>
        protected virtual DialogResults HandleResult(DialogResults result)
        {
            // We do need to check if the Microsoft Word dialog has been
            // canceled to avoid an exception on calling the method "Execute".
            if (!this.ResultIsOk(result))
            {
                try
                {
                    this.WordDialog.Execute();
                }
                catch (COMException ex)
                {
                    // We do need to handle the exception explicit when working
                    // with a built-in Microsoft Word dialog box.
                    MessageBox.Show(
                        ex.Message,
                        "Microsoft Word",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }

            return result;
        }

        /// <summary>
        /// Displays this <see cref="Dialog"/> until either the user closes it
        /// or the specified amount of time has passed.
        /// </summary>
        /// <param name="timeout">
        /// The amount of time that Word will wait before closing the dialog box
        /// automatically. One unit is approximately 0.001 second. Concurrent
        /// system activity may increase the effective time value. If this
        /// argument is omitted, the dialog box is closed when the user closes
        /// it.
        /// </param>
        /// <returns>
        /// An identifier of the enumeration <see cref="DialogResults"/>,
        /// indicating the return value of the built-in Microsoft Word dialog
        /// box.
        /// </returns>
        protected virtual DialogResults Display(int timeout = 0)
        {
            return (DialogResults)this.WordDialog.Display(timeout);
        }

        /// <summary>
        /// Checks whether the specified <see cref="DialogResults"/> is equal to
        /// <see cref="DialogResults.Cancel"/>.
        /// </summary>
        /// <param name="result">The <see cref="DialogResults"/> object to check.</param>
        /// <returns><c>true</c> on success; <c>false</c> on failure.</returns>
        protected bool ResultIsCancel(DialogResults result)
        {
            return result.Equals(DialogResults.Cancel);
        }

        /// <summary>
        /// Checks whether the specified <see cref="DialogResults"/> is equal to
        /// <see cref="DialogResults.Close"/>.
        /// </summary>
        /// <param name="result">The <see cref="DialogResults"/> object to check.</param>
        /// <returns><c>true</c> on success; <c>false</c> on failure.</returns>
        protected bool ResultIsClose(DialogResults result)
        {
            return result.Equals(DialogResults.Close);
        }

        /// <summary>
        /// Checks whether the specified <see cref="DialogResults"/> is equal to
        /// <see cref="DialogResults.Ok"/>.
        /// </summary>
        /// <param name="result">The <see cref="DialogResults"/> object to check.</param>
        /// <returns><c>true</c> on success; <c>false</c> on failure.</returns>
        protected bool ResultIsOk(DialogResults result)
        {
            return result.Equals(DialogResults.Ok);
        }
    }
}
