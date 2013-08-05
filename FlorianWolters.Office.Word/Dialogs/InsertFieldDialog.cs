//------------------------------------------------------------------------------
// <copyright file="InsertFieldDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using System.Windows.Forms;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="InsertFieldDialog"/> allows to interact with the
    /// built-in Microsoft Word dialog box "Insert field...".
    /// </summary>
    public class InsertFieldDialog : Dialog
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InsertFieldDialog"/>
        /// class with the specified <see cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        public InsertFieldDialog(Word.Application application)
            : base(application, Word.WdWordDialog.wdDialogInsertField)
        {
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
        protected override DialogResults HandleResult(DialogResults result)
        {
            if (this.ResultIsOk(result))
            {
                // TODO This isn't optimal, since we mix business logic with UI,
                // but since we do deal with built-in Microsoft Word dialogs
                // that isn't so bad. Further developments should abstract this
                // further.
                DialogResult dialogResult = MessageBox.Show(
                    "Do you really want to insert the field \""
                    + this.WordDialog.Field + "\" at the current position in the document?",
                    "Question",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                switch (dialogResult)
                {
                    case DialogResult.Yes:
                        base.HandleResult(result);
                        break;
                    case DialogResult.Cancel:
                        this.Show();
                        break;
                    default:
                        // NOOP
                        break;
                }
            }

            return result;
        }
    }
}
