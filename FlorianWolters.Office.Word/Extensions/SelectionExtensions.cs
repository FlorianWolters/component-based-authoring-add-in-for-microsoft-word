//------------------------------------------------------------------------------
// <copyright file="SelectionExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Extensions
{
    using System.Collections.Generic;
    using System.Linq;
    using Word = Microsoft.Office.Interop.Word;

    // TODO Refactor class.
    public static class SelectionExtensions
    {
        public static IEnumerable<Word.Field> SelectedFields(this Word.Selection selection)
        {
            // TODO Iterate paragraphs should be faster, but the current API does not seem to allow that.
            // IEnumerable<Field> fields = selection.Range.Paragraphs[1].;
            foreach (Word.Field field in selection.Document.Fields)
            {
                int fieldStart = field.Code.FormattedText.Start;
                string fieldResult = field.Result.Text;
                if (null == fieldResult)
                {
                    continue;
                }

                int displayedTextLength = fieldResult.Count();
                int fieldEnd = field.Code.FormattedText.End + displayedTextLength;

                if (!((fieldStart < selection.Start) & (fieldEnd < selection.Start)
                    | (fieldStart > selection.End) & (fieldEnd > selection.End)))
                {
                    yield return field;
                }
            }
        }

        public static IEnumerable<Word.Field> SelectedIncludeTextFields(this Word.Selection selection)
        {
            foreach (Word.Field field in selection.Document.Fields)
            {
                int fieldStart = field.Code.FormattedText.Start;
                string fieldResult = field.Result.Text;
                if (null == fieldResult)
                {
                    continue;
                }

                int displayedTextLength = fieldResult.Count();
                int fieldEnd = field.Code.FormattedText.End + displayedTextLength;

                if (!((fieldStart < selection.Start) & (fieldEnd < selection.Start)
                    | (fieldStart > selection.End) & (fieldEnd > selection.End))
                    && Word.WdFieldType.wdFieldIncludeText == field.Type)
                {
                    yield return field;
                }
            }
        }

        public static IEnumerable<Word.Field> SelectedIncludeFields(this Word.Selection selection)
        {
            foreach (Word.Field field in selection.Document.Fields)
            {
                int fieldStart = field.Code.FormattedText.Start;

                if (null != field.Result.Text)
                {
                    int displayedTextLength = field.Result.Text.Count();
                    int fieldEnd = field.Code.FormattedText.End + displayedTextLength;

                    if (!((fieldStart < selection.Start) & (fieldEnd < selection.Start)
                        | (fieldStart > selection.End) & (fieldEnd > selection.End))
                        && (Word.WdFieldType.wdFieldIncludeText == field.Type
                        || Word.WdFieldType.wdFieldIncludePicture == field.Type))
                    {
                        yield return field;
                    }
                }
            }
        }
    }
}
