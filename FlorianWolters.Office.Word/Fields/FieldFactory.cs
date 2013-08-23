//------------------------------------------------------------------------------
// <copyright file="FieldFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System;
    using System.Collections;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;
    using FlorianWolters.IO;
    using FlorianWolters.Office.Word.DocumentProperties;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="FieldFactory"/> simplifies the creation of new <see cref="Word.Field"/>s.
    /// </summary>
    public class FieldFactory
    {
        /// <summary>
        /// The <see cref="Word.Application"/> to interact with.
        /// </summary>
        private readonly Word.Application application;

        /// <summary>
        /// The <see cref="CustomDocumentPropertyReader"/> used to read custom document properties from the <see
        /// cref="Word.Document"/> of the <see cref="Word.Application"/> to interact with.
        /// </summary>
        private readonly CustomDocumentPropertyReader customDocumentPropertyReader;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFactory"/> class.
        /// </summary>
        /// <param name="application">The <see cref="Word.Application"/> to interact with.</param>
        public FieldFactory(Word.Application application)
            : this(application, new CustomDocumentPropertyReader())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldFactory"/> class.
        /// </summary>
        /// <param name="application">The <see cref="Word.Application"/> to interact with.</param>
        /// <param name="customDocumentPropertyReader">
        /// The <see cref="CustomDocumentPropertyReader"/> used to read custom document properties from the <see
        /// cref="Word.Document"/> of the <see cref="Word.Application"/> to interact with.
        /// </param>
        public FieldFactory(Word.Application application, CustomDocumentPropertyReader customDocumentPropertyReader)
        {
            this.application = application;
            this.customDocumentPropertyReader = customDocumentPropertyReader;
        }

        /// <summary>
        /// Adds a new <i>DATE</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertDate(Word.Range range, bool preserveFormatting = false)
        {
            return this.AddFieldToRange(range,  Word.WdFieldType.wdFieldDate, preserveFormatting);
        }

        /// <summary>
        /// Adds a new <i>DOCPROPERTY</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="propertyName">The name of a custom document property.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertDocProperty(Word.Range range, string propertyName, bool preserveFormatting = false)
        {
            this.customDocumentPropertyReader.Load(this.application.ActiveDocument);
            this.customDocumentPropertyReader.Get(propertyName);

            return this.AddFieldToRange(range, Word.WdFieldType.wdFieldDocProperty, preserveFormatting, propertyName);
        }

        /// <summary>
        /// Adds a new empty <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertEmpty(Word.Range range, bool preserveFormatting = false)
        {
            Word.Field result = this.AddFieldToRange(range, Word.WdFieldType.wdFieldEmpty, preserveFormatting);

            // Show the field codes of an empty field, because otherwise we can't be sure that it is visible.
            result.ShowCodes = true;

            return result;
        }

        /// <summary>
        /// Adds a new <i>LISTNUM</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertListNum(Word.Range range, bool preserveFormatting = false)
        {
            return this.AddFieldToRange(range, Word.WdFieldType.wdFieldListNum, preserveFormatting);
        }

        /// <summary>
        /// Adds a new <i>PAGE</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertPage(Word.Range range, bool preserveFormatting = false)
        {
            return this.AddFieldToRange(range, Word.WdFieldType.wdFieldPage, preserveFormatting);
        }

        /// <summary>
        /// Adds a new <i>TIME</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        public Word.Field InsertTime(Word.Range range, bool preserveFormatting = false)
        {
            return this.AddFieldToRange(range, Word.WdFieldType.wdFieldTime, preserveFormatting);
        }

        /// <summary>
        /// Adds a new <i>INCLUDEPICTURE</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="embed">Whether to store graphics data with the <see cref="Word.Document"/>.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        /// <exception cref="FileNotFoundException">If the specified file path does not exist.</exception>
        public Word.Field InsertIncludePicture(
            Word.Range range,
            string filePath,
            bool embed = false,
            bool preserveFormatting = true)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            StringBuilder text = this.CreateFileNameArgument(filePath);

            if (!embed)
            {
                text.Append(" \\d");
            }

            return this.AddFieldToRange(
                range,
                Word.WdFieldType.wdFieldIncludePicture,
                preserveFormatting,
                text.ToString());
        }

        /// <summary>
        /// Adds a new <i>INCLUDETEXT</i> <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="preventUpdatingOfFieldsInText">
        /// Whether to prevent updating fields in the inserted text unless the fields are first updated in the source
        /// document.
        /// </param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous <see cref="Word.Field"/> result to the new result.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        /// <exception cref="FileNotFoundException">If the specified file path does not exist.</exception>
        public Word.Field InsertIncludeText(
            Word.Range range,
            string filePath,
            bool preventUpdatingOfFieldsInText = false,
            bool preserveFormatting = false)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            StringBuilder text = this.CreateFileNameArgument(filePath);

            if (preventUpdatingOfFieldsInText)
            {
                text.Append(" \\!");
            }

            return this.AddFieldToRange(
                range,
                Word.WdFieldType.wdFieldIncludeText,
                preserveFormatting,
                text.ToString());
        }

        /// <summary>
        /// Adds a new <i>INCLUDEPICTURE</i> <see cref="Word.Field"/> which contains a nested <i>DOCPROPERTY</i>
        /// <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="propertyName">The name of a custom document property.</param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        /// <exception cref="FileNotFoundException">If the specified file path does not exist.</exception>
        public Word.Field InsertIncludePictureWithNestedDocProperty(
            Word.Range range,
            string filePath,
            string propertyName)
        {
            return this.InsertIncludeWithNestedDocProperty(range, filePath, propertyName, true);
        }

        /// <summary>
        /// Adds a new <i>INCLUDETEXT</i> <see cref="Word.Field"/> which contains a nested <i>DOCPROPERTY</i>
        /// <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="propertyName">The name of a custom document property.</param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        /// <exception cref="FileNotFoundException">If the specified file path does not exist.</exception>
        public Word.Field InsertIncludeTextWithNestedDocProperty(
            Word.Range range,
            string filePath,
            string propertyName)
        {
            return this.InsertIncludeWithNestedDocProperty(range, filePath, propertyName);
        }

        /// <summary>
        /// Adds one or more new <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// <para>
        /// This method allows to insert nested fields at the specified range.
        /// </para>
        /// <example>
        /// <c>InsertField(Application.Selection.Range, {{= {{PAGE}} - 1}};</c>
        /// will produce
        /// { = { PAGE } - 1 }
        /// </example>
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="theString">The string to convert to one or more <see cref="Word.Field"/> objects.</param>
        /// <param name="fieldOpen">The special code to mark the start of a <see cref="Word.Field"/>.</param>
        /// <param name="fieldClose">The special code to mark the end of a <see cref="Word.Field"/>.</param>
        /// <returns>The newly created <see cref="Word.Field"/></returns>
        /// <remarks>
        /// A solution for VBA has been taken from <a href="http://stoptyping.co.uk/word/nested-fields-in-vba">this</a>
        /// article and adopted for C# by the author.
        /// </remarks>
        public Word.Field InsertField(
            Word.Range range,
            string theString = "{{}}",
            string fieldOpen = "{{",
            string fieldClose = "}}")
        {
            if (null == range)
            {
                throw new ArgumentNullException("range");
            }

            if (string.IsNullOrEmpty(fieldOpen))
            {
                throw new ArgumentException("fieldOpen");
            }

            if (string.IsNullOrEmpty(fieldClose))
            {
                throw new ArgumentException("fieldClose");
            }

            if (!theString.Contains(fieldOpen) || !theString.Contains(fieldClose))
            {
                throw new ArgumentException("theString");
            }

            // Special case. If we do not check this, the algorithm breaks.
            if (theString == fieldOpen + fieldClose)
            {
                return this.InsertEmpty(range);
            }

            // TODO Implement additional error handling.

            // TODO Possible to remove the dependency to state capture?
            using (new StateCapture(range.Application.ActiveDocument))
            {
                Word.Field result = null;
                Stack fieldStack = new Stack();

                range.Text = theString;
                fieldStack.Push(range);

                Word.Range searchRange = range.Duplicate;
                Word.Range nextOpen = null;
                Word.Range nextClose = null;
                Word.Range fieldRange = null;

                while (searchRange.Start != searchRange.End)
                {
                    nextOpen = this.FindNextOpen(searchRange.Duplicate, fieldOpen);
                    nextClose = this.FindNextClose(searchRange.Duplicate, fieldClose);

                    if (null == nextClose)
                    {
                        break;
                    }

                    // See which marker comes first.
                    if (nextOpen.Start < nextClose.Start)
                    {
                        nextOpen.Text = string.Empty;
                        searchRange.Start = nextOpen.End;

                        // Field open, so push a new range to the stack.
                        fieldStack.Push(nextOpen.Duplicate);
                    }
                    else
                    {
                        nextClose.Text = string.Empty;

                        // Move start of main search region onwards past the end marker.
                        searchRange.Start = nextClose.End;

                        // Field close, so pop the last range from the stack and insert the field.
                        fieldRange = (Word.Range)fieldStack.Pop();
                        fieldRange.End = nextClose.End;
                        result = this.InsertEmpty(fieldRange);
                    }
                }

                // Move the current selection after all inserted fields.
                // TODO Improvement possible, e.g. by using another range object?
                int newPos = fieldRange.End + fieldRange.Fields.Count + 1;
                fieldRange.SetRange(newPos, newPos);
                fieldRange.Select();

                // Update the result of the outer field object.
                result.Update();

                return result;
            }
        }

        /// <summary>
        /// Returns a <see cref="Word.Range"/> which contains the next start of a field.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> to search in.</param>
        /// <param name="text">The text to be searched for.</param>
        /// <returns>A new <see cref="Word.Range"/> which contains the next start of a field.</returns>
        private Word.Range FindNextOpen(Word.Range range, string text)
        {
            Word.Find find = this.CreateFind(range, text);
            Word.Range result = range.Duplicate;

            if (!find.Found)
            {
                // Make sure that the next closing field will be found first.
                result.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }

            return result;
        }

        /// <summary>
        /// Returns a <see cref="Word.Range"/> which contains the next end of a field.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> to search in.</param>
        /// <param name="text">The text to be searched for.</param>
        /// <returns>A new <see cref="Word.Range"/> which contains the next end of a field.</returns>
        private Word.Range FindNextClose(Word.Range range, string text)
        {
            return this.CreateFind(range, text).Found ? range.Duplicate : null;
        }

        /// <summary>
        /// Modifies the specified <see cref="Word.Range"/> to match the text to be searched for.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> to search in and to modify.</param>
        /// <param name="text">The text to be searched for.</param>
        /// <returns>The modified <see cref="Word.Range"/>.</returns>
        private Word.Find CreateFind(Word.Range range, string text)
        {
            Word.Find result = range.Find;
            result.Execute(FindText: text, Forward: true, Wrap: Word.WdFindWrap.wdFindStop);

            return result;
        }

        /// <summary>
        /// Creates a <see cref="Word.Field"/> and adds it to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <remarks>
        /// The <see cref="Word.Field"/> is added to the <see cref="Word.Fields"/> collection of the specified <see
        /// cref="Word.Range"/>.
        /// </remarks>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="type">The type of <see cref="Word.Field"/> to create.</param>
        /// <param name="preserveFormatting">
        /// Whether to apply the formatting of the previous field result to the new result.
        /// </param>
        /// <param name="text">Additional text needed for the <see cref="Word.Field"/>.</param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        private Word.Field AddFieldToRange(
            Word.Range range,
            Word.WdFieldType type,
            bool preserveFormatting = false,
            string text = null)
        {
            try
            {
                return range.Fields.Add(
                    range,
                    type,
                    (null == text) ? Type.Missing : text,
                    preserveFormatting);
            }
            catch (COMException)
            {
                throw new FieldCreationException("Unable to create a field at the current selection.");
            }
        }

        /// <summary>
        /// Creates the file name argument for a <i>INCLUDE[...]</i> field.
        /// </summary>
        /// <param name="filePath">The file path to use.</param>
        /// <returns>The file name argument.</returns>
        private StringBuilder CreateFileNameArgument(string filePath)
        {
            StringBuilder result = new StringBuilder("\"");
            result.Append(new FieldFilePathTranslator().Encode(filePath));
            result.Append("\"");

            return result;
        }

        /// <summary>
        /// Adds a new <i>INCLUDE[...]</i> <see cref="Word.Field"/> which contains a nested <i>DOCPROPERTY</i>
        /// <see cref="Word.Field"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> where to add the <see cref="Word.Field"/>.</param>
        /// <param name="filePath">The file path of the file to include.</param>
        /// <param name="propertyName">The name of a custom document property.</param>
        /// <param name="includePicture">
        /// Whether the <see cref="Word.Field"/> to create is of type <i>IncludePicture</i>.
        /// </param>
        /// <returns>The newly created <see cref="Word.Field"/>.</returns>
        /// <exception cref="FileNotFoundException">If the specified file path does not exist.</exception>
        private Word.Field InsertIncludeWithNestedDocProperty(
            Word.Range range,
            string filePath,
            string propertyName,
            bool includePicture = false)
        {
            this.ThrowFileNotFoundExceptionIfFileDoesNotExist(filePath);

            string documentFilePath = range.Application.ActiveDocument.Path;
            string relativeFilePath = filePath;

            if (Path.IsPathRooted(filePath))
            {
                // Convert the absolute file path to a relative file path.
                relativeFilePath = PathUtils.GetRelativePath(documentFilePath, filePath);
            }

            StringBuilder fieldText = new StringBuilder("{{");

            if (includePicture)
            {
                fieldText.Append("INCLUDEPICTURE");
            }
            else
            {
                fieldText.Append("INCLUDETEXT");
            }

            fieldText.Append(" \"{{DOCPROPERTY ");
            fieldText.Append(propertyName);
            fieldText.Append("}}\\\\");
            fieldText.Append(new FieldFilePathTranslator().Encode(relativeFilePath));
            fieldText.Append("\"");

            if (includePicture)
            {
                fieldText.Append(" \\d \\* MERGEFORMAT");
            }

            fieldText.Append("}}");

            return this.InsertField(range, fieldText.ToString());
        }

        /// <summary>
        /// Throws a new <see cref="FileNotFoundException"/> if the specified file path does not exist.
        /// </summary>
        /// <param name="filePath">The file path to check.</param>
        /// <exception cref="FileNotFoundException">If the specified file path does not exist.</exception>
        private void ThrowFileNotFoundExceptionIfFileDoesNotExist(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The file \"" + filePath + "\" does not exist.");
            }
        }
    }
}
