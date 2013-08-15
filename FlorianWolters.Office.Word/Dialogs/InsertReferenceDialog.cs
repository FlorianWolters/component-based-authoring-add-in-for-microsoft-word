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

    /// <summary>
    /// The abstract class <see cref="InsertReferenceDialog"/> allows to
    /// interact with a built-in Microsoft Word dialog box that is used to
    /// insert a reference to another file.
    /// </summary>
    public abstract class InsertReferenceDialog : Dialog
    {
        /// <summary>
        /// The <see cref="FieldFactory"/> used to create fields.
        /// </summary>
        protected readonly FieldFactory FieldFactory;

        /// <summary>
        /// The name of the custom document property which contains the absolute
        /// directory path of the active document.
        /// </summary>
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
                if (!this.Application.ActiveDocument.IsSaved())
                {
                    MessageBox.Show(
                        "The target document (this document) has to be saved first, to be able to include a reference to a source file.",
                        "Attention",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
                }
                else
                {
                    try
                    {
                        this.CreateField();

                        if (!this.HasUserChosenLinkToFile())
                        {
                            MessageBox.Show(
                                "A source file has to be linked to be reusable. Therefore the selected source file to include has been converted to a link.",
                                "Information",
                                 MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                    catch (ArgumentException)
                    {
                        MessageBox.Show(
                            "The source document (the document to insert) has to be stored on the identical drive as the target document (this document).",
                            "Attention",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Inserts a <see cref="Word.Field"/> into the current <see
        /// cref="Word.Selection"/> of the active <see cref="Word.Document"/>.
        /// <para>
        /// The field can be an <i>INCLUDETEXT</i> or <i>INCLUDEPICTURE</i>
        /// field, for example.
        /// </para>
        /// </summary>
        protected abstract void CreateField();

        /// <summary>
        /// Determines whether the user has chosen to create a reference or not.
        /// </summary>
        /// <returns>
        /// <c>true</c> if the user has chosen to create a reference or
        /// <c>false</c> if not.
        /// </returns>
        protected abstract bool HasUserChosenLinkToFile();

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
    }
}
