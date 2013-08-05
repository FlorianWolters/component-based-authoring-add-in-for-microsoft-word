//------------------------------------------------------------------------------
// <copyright file="FieldFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using FlorianWolters.IO;
    using FlorianWolters.Office.Word.DocumentProperties;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="FieldFactory"/> simplifies the creation of new <see
    /// cref="Word.Field"/>s.
    /// </summary>
    public class FieldFactory
    {
        /// <summary>
        /// The <see cref="Word.Application"/> to interact with.
        /// </summary>
        private readonly Word.Application application;

        /// <summary>
        /// The <see cref="CustomDocumentPropertyReader"/> used to read custom
        /// document properties from the <see cref="Word.Document"/> of the <see
        /// cref="Word.Application"/> to interact with.
        /// </summary>
        private readonly CustomDocumentPropertyReader customDocumentPropertyReader;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFactory"/> class.
        /// </summary>
        /// <param name="application">
        /// The <see cref="Word.Application"/> to interact with.
        /// </param>
        public FieldFactory(Word.Application application)
            : this(application, new CustomDocumentPropertyReader())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFactory"/> class.
        /// </summary>
        /// <param name="application">
        /// The <see cref="Word.Application"/> to interact with.
        /// </param>
        /// <param name="customDocumentPropertyReader">
        /// The <see cref="CustomDocumentPropertyReader"/> used to read custom
        /// document properties from the <see cref="Word.Document"/> of the <see
        /// cref="Word.Application"/> to interact with.
        /// </param>
        public FieldFactory(
            Word.Application application,
            CustomDocumentPropertyReader customDocumentPropertyReader)
        {
            this.application = application;
            this.customDocumentPropertyReader = customDocumentPropertyReader;
        }

        /// <summary>
        /// Inserts a new <i>DATE</i> <see cref="Word.Field"/> into the current
        /// <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertDate(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldDate,
                mergeFormat);
        }

        /// <summary>
        /// Inserts a new <i>DOCPROPERTY</i> <see cref="Word.Field"/> into the
        /// current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="propertyName">
        /// The name of a custom document property.
        /// </param>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertDocProperty(
            string propertyName,
            bool mergeFormat = false)
        {
            this.customDocumentPropertyReader.Load(
                this.application.ActiveDocument);
            this.customDocumentPropertyReader.Get(propertyName);

            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldDocProperty,
                mergeFormat,
                propertyName);
        }

        /// <summary>
        /// Inserts a new empty <see cref="Word.Field"/> into the current <see
        /// cref="Word.Selection"/>.
        /// </summary>
        /// <param name="data">The optional data of the field.</param>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertEmpty(
            string data = null,
            bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldEmpty,
                mergeFormat,
                data);
        }

        /// <summary>
        /// Inserts a new <i>LISTNUM</i> <see cref="Word.Field"/> into the
        /// current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertListNum(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldListNum,
                mergeFormat);
        }

        /// <summary>
        /// Inserts a new <i>PAGE</i> <see cref="Word.Field"/> into the current
        /// <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertPage(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldPage,
                mergeFormat);
        }

        /// <summary>
        /// Inserts a new <i>TIME</i> <see cref="Word.Field"/> into the current
        /// <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertTime(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldTime,
                mergeFormat);
        }

        /// <summary>
        /// Inserts a new <i>INCLUDEPICTURE</i> <see cref="Word.Field"/> into
        /// the current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertIncludePicture(
            string filePath,
            bool mergeFormat = false)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldIncludePicture,
                mergeFormat,
                filePath);
        }

        /// <summary>
        /// Inserts a new <i>INCLUDETEXT</i> <see cref="Word.Field"/> into the
        /// current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertIncludeText(
            string filePath,
            bool mergeFormat = false)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldIncludeText,
                mergeFormat,
                filePath);
        }

        /// <summary>
        /// Inserts a new <i>INCLUDEPICTURE</i> <see cref="Word.Field"/> which
        /// contains a nested <i>DOCPROPERTY</i> <see cref="Word.Field"/> into
        /// the current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="propertyName">
        /// The name of the nested <i>DOCPROPERTY</i> field.
        /// </param>
        public void InsertIncludePictureWithNestedDocProperty(
            string filePath,
            string propertyName)
        {
            this.InsertIncludeFieldWithNestedDocProperty(
                filePath,
                propertyName,
                "INCLUDEPICTURE");
            
            // Append the "\d" switch to avoid that the graphics data is saved
            // in the Word document.
            this.application.Selection.Range.InsertAfter("\\d");
        }

        /// <summary>
        /// Inserts a new <i>INCLUDETEST</i> <see cref="Word.Field"/> which
        /// contains a nested <i>DOCPROPERTY</i> <see cref="Word.Field"/> into
        /// the current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="propertyName">
        /// The name of the nested <i>DOCPROPERTY</i> field.
        /// </param>
        public void InsertIncludeTextWithNestedDocProperty(
            string filePath,
            string propertyName)
        {
            this.InsertIncludeFieldWithNestedDocProperty(
                filePath,
                propertyName,
                "INCLUDETEXT");
        }

        /// <summary>
        /// Adds the specified type of <see cref="Word.Field"/> to the current
        /// <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="type">
        /// The type of <see cref="Word.Field"/> to create.
        /// </param>
        /// <param name="mergeFormat">
        /// Whether to apply the formatting of the previous field result to the
        /// new result.
        /// </param>
        /// <param name="text">The text of the <see cref="Word.Field"/>.</param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        private Word.Field AddFieldToCurrentSelection(
            Word.WdFieldType type,
            bool mergeFormat = false,
            string text = null)
        {
            Word.Field result = null;
            Word.Selection selection = this.application.Selection;
            Word.Range range = selection.Range;

            try
            {
                result = range.Fields.Add(
                    range,
                    type,
                    text,
                    mergeFormat);
            }
            catch (COMException)
            {
                throw new FieldCreationException(
                    "Unable to create a field at the current selection.");
            }

            return result;
        }

        /// <summary>
        /// Throws a new <see cref="FileNotFoundException"/> if the specified
        /// file path does not exist.
        /// </summary>
        /// <param name="filePath">The file path to check.</param>
        /// <exception cref="FileNotFoundException">
        /// If the specified file path does not exist.
        /// </exception>
        private void ThrowFileNotFoundExceptionIfFileDoesNotExist(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(
                    "The file \"" + filePath + "\" does not exist.");
            }
        }

        // TODO Find a better solution.
        // http://stackoverflow.com/questions/16774411/create-a-nested-field-with-visual-studio-tools-for-office-vsto
       
        /// <summary>
        /// Inserts a new <i>INCLUDE[...]</i> <see cref="Word.Field"/> into
        /// the current <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="propertyName">
        /// The name of the nested <i>DOCPROPERTY</i> field.
        /// </param>
        /// <param name="functionName">
        /// The name of the <see cref="Word.Field"/> function, e.g.
        /// <c>INCLUDETEXT</c>.
        /// </param>
        private void InsertIncludeFieldWithNestedDocProperty(
            string filePath,
            string propertyName,
            string functionName)
        {
            string documentFilePath = this.application.ActiveDocument.Path;
            string absoluteFilePath = filePath;
            string relativeFilePath = filePath;

            if (Path.IsPathRooted(filePath))
            {
                // Convert a possible absolute file path to a relative path
                relativeFilePath = PathUtils.GetRelativePath(
                    documentFilePath,
                    filePath);
            }
            else
            {
                absoluteFilePath = documentFilePath + Path.DirectorySeparatorChar + relativeFilePath;
                this.ThrowFileNotFoundExceptionIfFileDoesNotExist(absoluteFilePath);
            }

            DateTime lastModified = File.GetLastWriteTimeUtc(absoluteFilePath);

            this.application.ScreenUpdating = false;
            this.application.ActiveWindow.View.ShowFieldCodes = true;

            Word.Selection selection = this.application.ActiveWindow.Selection;
            this.InsertDocProperty(propertyName);

            // Select the previously inserted DocProperty field.
            selection.MoveLeft(
                Unit: Word.WdUnits.wdWord,
                Count: 1,
                Extend: Word.WdMovementType.wdExtend);

            // Create a new empty field AROUND the DocProperty field. After that
            // the DocProperty field is nested INSIDE the empty field.
            selection.Range.Fields.Add(
                selection.Range,
                Word.WdFieldType.wdFieldEmpty,
                PreserveFormatting: false);

            selection.InsertAfter(functionName + " \"");

            // Move the selection AFTER the inner field.
            selection.MoveRight(
                Unit: Word.WdUnits.wdWord,
                Count: 1,
                Extend: Word.WdMovementType.wdExtend);

            // Insert text AFTER the nested field.
            selection.InsertAfter(
                "\\\\" + new FieldFilePathTranslator().Encode(relativeFilePath) + "\"");

            selection.MoveRight(
                Unit: Word.WdUnits.wdWord,
                Count: 1,
                Extend: Word.WdMovementType.wdMove);

            // Insert the last modified datetime of the reference file AFTER the INCLUDE field.
            selection.Range.Fields.Add(
                selection.Range,
                Word.WdFieldType.wdFieldEmpty,
                Text: lastModified.ToString("u"),
                PreserveFormatting: false);

            selection.Fields.Update();
            this.application.ActiveWindow.View.ShowFieldCodes = false;
            this.application.ScreenUpdating = true;
        }
    }
}
