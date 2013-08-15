//------------------------------------------------------------------------------
// <copyright file="ExtendedIncludeFieldComparison.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using Word = Microsoft.Office.Interop.Word;

    public class ExtendedIncludeFieldComparison
    {
        private readonly Word.Application application;

        public ExtendedIncludeFieldComparison(Word.Application application)
        {
            this.application = application;
        }

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

            // The result from the INCLUDE field is assumed to be the revised (temporary) document by convention.
            Word.Document revisedDocument = this.application.Documents.Add(
                Template: template.FullName,
                DocumentType: Word.WdNewDocumentType.wdNewBlankDocument,
                Visible: false);

            // Copy the result of the field to the original (temporary document).
            field.Result.TextRetrievalMode.IncludeFieldCodes = false;
            field.Result.TextRetrievalMode.IncludeHiddenText = true;

            // TODO The content of has a trailing '\r' character too much. I don't know why, since the field result is
            // correct and the Text property should replace the complete text of the range.
            revisedDocument.Content.Text = field.Result.Text.TrimEnd(new char[] { '\r' });

            // Run the comparison.
            Word.Document diffDocument = this.application.CompareDocuments(originalDocument, revisedDocument);

            // Modify the UI for the "diff" document.
            diffDocument.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneRevisionsVert;
            diffDocument.ActiveWindow.ShowSourceDocuments = Word.WdShowSourceDocuments.wdShowSourceDocumentsBoth;

            // Close the revised document and the original  document.
            ((Word._Document)revisedDocument).Close(SaveChanges: false);
            ((Word._Document)originalDocument).Close(SaveChanges: false);

            return diffDocument;
        }
    }
}
