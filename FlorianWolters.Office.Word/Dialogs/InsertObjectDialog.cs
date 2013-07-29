//------------------------------------------------------------------------------
// <copyright file="InsertObjectDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using Word = Microsoft.Office.Interop.Word;

    public class InsertObjectDialog : Dialog
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InsertObjectDialog"/>
        /// class with the specified <see cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        public InsertObjectDialog(Word.Application application)
            : base(application, Word.WdWordDialog.wdDialogInsertObject)
        {
        }
    }
}
