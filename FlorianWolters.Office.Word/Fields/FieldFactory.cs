//------------------------------------------------------------------------------
// <copyright file="FieldFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System.IO;
    using System.Runtime.InteropServices;
    using FlorianWolters.IO;
    using FlorianWolters.Office.Word.DocumentProperties;
    using Word = Microsoft.Office.Interop.Word;

    public class FieldFactory
    {
        private readonly Word.Application application;
        private readonly CustomDocumentPropertyReader customDocumentPropertyReader;

        public FieldFactory(Word.Application application)
            : this(application, new CustomDocumentPropertyReader())
        {
        }

        public FieldFactory(
            Word.Application application,
            CustomDocumentPropertyReader customDocumentPropertyReader)
        {
            this.application = application;
            this.customDocumentPropertyReader = customDocumentPropertyReader;
        }

        public Word.Field InsertDate(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldDate,
                mergeFormat);
        }

        public Word.Field InsertDocProperty(string propertyName, bool mergeFormat = false)
        {
            this.customDocumentPropertyReader.Load(this.application.ActiveDocument);
            this.customDocumentPropertyReader.Get(propertyName);

            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldDocProperty,
                mergeFormat,
                propertyName);
        }

        public Word.Field InsertEmpty(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldEmpty,
                mergeFormat,
                this.application.Selection.Range.Text);
        }

        public Word.Field InsertListNum(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldListNum,
                mergeFormat);
        }

        public Word.Field InsertPage(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldPage,
                mergeFormat);
        }

        public Word.Field InsertTime(bool mergeFormat = false)
        {
            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldTime,
                mergeFormat);
        }

        public Word.Field InsertIncludePicture(string filePath, bool mergeFormat = false)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldIncludePicture,
                mergeFormat,
                filePath);
        }

        public Word.Field InsertIncludeText(string filePath, bool mergeFormat = false)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            return this.AddFieldToCurrentSelection(
                Word.WdFieldType.wdFieldIncludeText,
                mergeFormat,
                filePath);
        }

        public void InsertIncludePictureWithNestedDocProperty(string filePath, string propertyName)
        {
            this.InsertIncludeFieldWithNestedDocProperty(filePath, propertyName, "INCLUDEPICTURE");
        }

        public void InsertIncludeTextWithNestedDocProperty(string filePath, string propertyName)
        {
            this.InsertIncludeFieldWithNestedDocProperty(filePath, propertyName, "INCLUDETEXT");
        }

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
        private void InsertIncludeFieldWithNestedDocProperty(
            string filePath,
            string propertyName,
            string functionName)
        {
            string documentFilePath = this.application.ActiveDocument.Path;

            if (Path.IsPathRooted(filePath))
            {
                // Convert a possible absolute file path to a relative path
                filePath = PathUtils.GetRelativePath(
                    documentFilePath,
                    filePath);
            }
            else
            {
                this.ThrowFileNotFoundExceptionIfFileDoesNotExist(
                    documentFilePath + Path.DirectorySeparatorChar + filePath);
            }

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
                "\\\\" + new FieldFilePathTranslator().Encode(filePath) + "\"");
            selection.Fields.Update();

            this.application.ActiveWindow.View.ShowFieldCodes = false;
            this.application.ScreenUpdating = true;
        }
    }
}
