//------------------------------------------------------------------------------
// <copyright file="InsertFileDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using System.IO;
    using FlorianWolters.Office.Word.Fields;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="InsertFileDialog"/> allows to interact with the
    /// built-in Microsoft Word dialog box "Insert file".
    /// </summary>
    public class InsertFileDialog : InsertReferenceDialog
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InsertFileDialog"/>
        /// class with the specified <see cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        /// <param name="fieldFactory">Required to create fields.</param>
        /// <param name="customDocumentPropertyNameWithDocumentPath">
        /// The name of the custom document property that contains the absolute
        /// directory path of the active Microsoft Word document.
        /// </param>
        public InsertFileDialog(
            Word.Application application,
            FieldFactory fieldFactory,
            string customDocumentPropertyNameWithDocumentPath) : base(
                application,
                Word.WdWordDialog.wdDialogInsertFile,
                fieldFactory,
                customDocumentPropertyNameWithDocumentPath)
        {
        }

        /// <summary>
        /// Inserts a <see cref="Word.Field"/> into the current <see
        /// cref="Word.Selection"/> of the active <see cref="Word.Document"/>.
        /// </summary>
        protected override void CreateField()
        {
            this.FieldFactory.InsertIncludeTextWithNestedDocProperty(
                this.AbsoluteFilePathOfTargetFile(),
                this.CustomDocumentPropertyNameWithDocumentPath);
        }

        /// <summary>
        /// Retrieves the absolute directory path of the active Microsoft Word
        /// document.
        /// </summary>
        /// <returns>The absolute directory path.</returns>
        protected override string AbsoluteFilePathOfTargetFile()
        {
            // We can't detect whether the user has chosen to link to the file
            // or not. Therefore we do have to identify the decision by
            // analyzing the return value of the built-in Microsoft Word dialog
            // box. If the return value specifies a file path the user has
            // chosen to link to the file, otherwise the return value specifies
            // a file propertyName and we have to built the file path on our own.
            string result = this.WordDialog.Name.Trim('"');

            if (!this.HasUserChosenLinkToFile())
            {
                string directoryPath = this.Application.Options.DefaultFilePath[Word.WdDefaultFilePath.wdDocumentsPath];
                result = directoryPath + Path.DirectorySeparatorChar + result;
            }

            return result;
        }

        /// <summary>
        /// Determines whether the user has chosen to create a reference or not.
        /// </summary>
        /// <returns>
        /// <c>true</c> if the user has chosen to create a reference or
        /// <c>false</c> if not.
        /// </returns>
        protected override bool HasUserChosenLinkToFile()
        {
            return Path.IsPathRooted(this.WordDialog.Name.Trim('"'));
        }
    }
}
