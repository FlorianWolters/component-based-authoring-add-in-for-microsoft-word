//------------------------------------------------------------------------------
// <copyright file="ExtendedIncludeFieldComparison.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="ExtendedIncludeFieldComparison"/> allows to compare the result of a <see
    /// cref="ExtendedIncludeField"/> with the content of its source file.
    /// </summary>
    public class ExtendedIncludeFieldComparison
    {
        /// <summary>
        /// The new line character used in a <see cref="Word.Document"/>.
        /// </summary>
        private const char WordDocumentNewLine = '\r';

        /// <summary>
        /// The <see cref="Word.Application"/> in which to execute the comparison.
        /// </summary>
        private readonly Word.Application application;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedIncludeFieldComparison"/> class.
        /// </summary>
        /// <param name="application">The <see cref="Word.Application"/> in which to execute the comparison.</param>
        public ExtendedIncludeFieldComparison(Word.Application application)
        {
            this.application = application;
        }

        /// <summary>
        /// Executes the comparison.
        /// </summary>
        /// <param name="extendedIncludeField">
        /// Contains the <see cref="Word.Field"/> result and the original file name.
        /// </param>
        /// <returns>
        /// A newly created <see cref="Word.Document"/> that contains the differences between the result of the field
        /// and the content of the source file, marked using tracked changes.
        /// </returns>
        public Word.Document Execute(ExtendedIncludeField extendedIncludeField)
        {
            Word.Field field = extendedIncludeField.IncludeField;

            // The source document is assumed to be the original document.
            Word.Document originalDocument = this.application.Documents.Open(
                FileName: extendedIncludeField.FilePath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);

            // Retrieve the template of the current document to attach it to the temporary document.
            Word.Template template = (Word.Template)this.application.ActiveDocument.get_AttachedTemplate();

            // TODO Poor workaround that creates another (temporary) document to allow comparison with nested INCLUDE
            // fields. The property Range.Text automatically resolves all fields. Via this method we do compare text
            // only.
            Word.Document originalDocumentTemp = this.application.Documents.Add(
                Template: template.FullName,
                DocumentType: Word.WdNewDocumentType.wdNewBlankDocument,
                Visible: false);
            originalDocumentTemp.Content.Text = this.NormalizeRangeText(originalDocument.Content);

            // The result from the INCLUDE field is assumed to be the revised (temporary) document by convention.
            Word.Document revisedDocument = this.application.Documents.Add(
                Template: template.FullName,
                DocumentType: Word.WdNewDocumentType.wdNewBlankDocument,
                Visible: false);

            // Copy the result of the field to the original (temporary document).
            Word.Range fieldResultRange = field.Result;
            fieldResultRange.TextRetrievalMode.IncludeFieldCodes = false;
            fieldResultRange.TextRetrievalMode.IncludeHiddenText = true;
            revisedDocument.Content.Text = this.NormalizeRangeText(fieldResultRange);

            // Run the comparison.
            Word.Document diffDocument = this.application.CompareDocuments(originalDocumentTemp, revisedDocument);

            // Modify the UI for the "diff" document.
            diffDocument.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneRevisionsVert;
            diffDocument.ActiveWindow.ShowSourceDocuments = Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth;

            // Close the revised document and the original  document.
            ((Word._Document)revisedDocument).Close(SaveChanges: false);
            ((Word._Document)originalDocument).Close(SaveChanges: false);
            ((Word._Document)originalDocumentTemp).Close(SaveChanges: false);

            return diffDocument;
        }

        /// <summary>
        /// Normalizes the text of the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> whose text to normalize.</param>
        /// <returns>The normalized text.</returns>
        private string NormalizeRangeText(Word.Range range)
        {
            // TODO The tetx of the range objects has a trailing '\r' character too much. I don't know why, since the
            // debugger shows the correct string and the Text property should replace the complete text of the range.
            return range.Text.TrimEnd(new char[] { WordDocumentNewLine });
        }
    }
}
