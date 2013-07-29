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

        protected override DialogResults HandleResult(DialogResults result)
        {
            if (this.ResultIsOK(result))
            {
                // TODO This isn't optimal, since we mix business logic with UI,
                // but since we do deal with built-in Microsoft Word dialogs
                // that isn't so bad. Further developments should abstract this
                // further.
                switch (this.ShowQuestion())
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

        private DialogResult ShowQuestion()
        {
            return MessageBox.Show(
                "Do you really want to insert the field \""
                + this.WordDialog.Field + "\" at the current position in the Document?",
                "Question",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);
        }
    }
}
