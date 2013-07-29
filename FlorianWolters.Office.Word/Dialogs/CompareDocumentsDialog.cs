//------------------------------------------------------------------------------
// <copyright file="CompareDocumentsDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using Word = Microsoft.Office.Interop.Word;

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
    }
}
