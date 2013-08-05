//------------------------------------------------------------------------------
// <copyright file="CompareDocumentsDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="CompareDocumentsDialog"/> allows to interact with
    /// the built-in Microsoft Word dialog box "Compare documents".
    /// </summary>
    public class CompareDocumentsDialog : Dialog
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="CompareDocumentsDialog"/> class with the specified <see
        /// cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        public CompareDocumentsDialog(Word.Application application)
            : base(application, Word.WdWordDialog.wdDialogToolsCompareDocuments)
        {
        }

        /// <summary>
        /// Displays and carries out actions initiated in the specified built-in
        /// Microsoft Word dialog box represented by this <see cref="Dialog"/>.
        /// </summary>
        /// <returns>
        /// An identifier of the enumeration <see cref="DialogResults"/>,
        /// indicating the return value of this <see cref="Dialog"/>.
        /// </returns>
        public override DialogResults Show()
        {
            // We do need to overwrite this method, since
            // Word.WdWordDialog.wdDialogToolsCompareDocuments always returns
            // DialogResults.Ok, no matter which button has been clicked to
            // close the dialog.
            return (DialogResults)this.WordDialog.Show();
        }
    }
}
