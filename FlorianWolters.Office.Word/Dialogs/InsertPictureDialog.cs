//------------------------------------------------------------------------------
// <copyright file="InsertPictureDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using FlorianWolters.Office.Word.Fields;
    using Word = Microsoft.Office.Interop.Word;

    public class InsertPictureDialog : InsertReferenceDialog
    {
        /// <summary>
        /// <i>Magic Number</i> which signals that the user has chosen to link
        /// to the picture to insert.
        /// </summary>
        private const int LinkToFileEnabled = 2;

        /// <summary>
        /// Initializes a new instance of the <see cref="InsertPictureDialog"/>
        /// class with the specified <see cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        /// <param name="fieldFactory">Required to create fields.</param>
        /// <param name="customDocumentPropertyNameWithDocumentPath">
        /// The name of the custom document property that contains the absolute
        /// directory path of the active Microsoft Word document.
        /// </param>
        public InsertPictureDialog(
            Word.Application application,
            FieldFactory fieldFactory,
            string customDocumentPropertyNameWithDocumentPath) : base(
                application,
                Word.WdWordDialog.wdDialogInsertPicture,
                fieldFactory,
                customDocumentPropertyNameWithDocumentPath)
        {
        }

        protected override void CreateField()
        {
            this.FieldFactory.InsertIncludePictureWithNestedDocProperty(
                this.AbsoluteFilePathOfTargetFile(),
                this.CustomDocumentPropertyNameWithDocumentPath);
        }

        protected override bool HasUserChosenLinkToFile()
        {
            return LinkToFileEnabled == int.Parse(this.WordDialog.LinkToFile);
        }
    }
}
