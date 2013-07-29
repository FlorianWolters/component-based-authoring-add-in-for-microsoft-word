//------------------------------------------------------------------------------
// <copyright file="InsertReferenceDialog.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    using System;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using Word = Microsoft.Office.Interop.Word;

    public abstract class InsertReferenceDialog : Dialog
    {
        protected readonly FieldFactory FieldFactory;
        protected readonly string CustomDocumentPropertyNameWithDocumentPath;

        /// <summary>
        /// Initializes a new instance of the <see cref="InsertReferenceDialog"/> class 
        /// with the specified <see cref="Word.Application"/> and the specified 
        /// <see cref="Word.WdWordDialog"/>.
        /// </summary>
        /// <param name="application">The Microsoft Word application.</param>
        /// <param name="dialogType">
        /// The type of the built-in Microsoft Word dialog box.
        /// </param>
        /// <param name="fieldFactory">Required to create fields.</param>
        /// <param name="customDocumentPropertyNameWithDocumentPath">
        /// The name of the custom document property that contains the absolute
        /// directory path of the active Microsoft Word document.
        /// </param>
        protected InsertReferenceDialog(
            Word.Application application,
            Word.WdWordDialog dialogType,
            FieldFactory fieldFactory,
            string customDocumentPropertyNameWithDocumentPath)
            : base(application, dialogType)
        {
            this.FieldFactory = fieldFactory;
            this.CustomDocumentPropertyNameWithDocumentPath = customDocumentPropertyNameWithDocumentPath;
        }

        protected abstract void CreateField();

        protected abstract bool HasUserChosenLinkToFile();

        protected override DialogResults HandleResult(DialogResults result)
        {
            if (this.ResultIsOk(result))
            {
                if (!this.Application.ActiveDocument.IsSaved())
                {
                    this.ShowMessageSaveActiveDocumentFirst();
                }
                else
                {
                    try
                    {
                        this.CreateField();

                        if (!this.HasUserChosenLinkToFile())
                        {
                            this.ShowMessageEnableLinkToFile();
                        }
                    }
                    catch (ArgumentException)
                    {
                        this.ShowMessageFileIdenticalDrive();
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Retrieves the absolute directory path of the active Microsoft Word
        /// document.
        /// </summary>
        /// <returns>The absolute directory path.</returns>
        protected string AbsoluteDirectoryPathOfSourceFile()
        {
            return this.Application.ActiveDocument.Path;
        }

        /// <summary>
        /// Retrieves the absolute file path of the file to insert into the
        /// active Microsoft Word document.
        /// </summary>
        /// <returns>The absolute file path.</returns>
        protected virtual string AbsoluteFilePathOfTargetFile()
        {
            return this.WordDialog.Name;
        }

        protected void ShowMessageEnableLinkToFile()
        {
            MessageBox.Show(
                "A source file has to be linked to be reusable. Therefore the selected source file to include has been converted to a link.",
                "Information",
                 MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ShowMessageSaveActiveDocumentFirst()
        {
            MessageBox.Show(
                "The target Document (this Document) has to be saved first, to be able to include a reference to a source file.",
                "Attention",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
        }

        private void ShowMessageFileIdenticalDrive()
        {
            MessageBox.Show(
                "The source Document (the Document to insert) has to be stored on the identical drive as the target Document (this Document).",
                "Attention",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
        }
    }
}
