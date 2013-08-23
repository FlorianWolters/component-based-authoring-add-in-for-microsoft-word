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

    /// <summary>
    /// The static class <see cref="SelectionExtensions"/> contains extension methods for a selection in a Microsoft
    /// Word document, represented by an object of the class <see cref="Word.Selection"/>.
    /// </summary>
    public static class SelectionExtensions
    {
        /// <summary>
        /// Retrieves all <see cref="Word.Field"/>s in the specified <see cref="Word.Selection"/>.
        /// <para>
        /// The property <c>Fields</c> in <see cref="Word.Selection"/> does only include a <see cref="Word.Field"/> if
        /// it is completely part of the <see cref="Word.Selection"/>. In contrast to that, this method does also
        /// include a <see cref="Word.Field"/> if it is only partially part of the <see cref="Word.Selection"/>.
        /// </para>
        /// </summary>
        /// <param name="selection">The <see cref="Word.Selection"/> to check for <see cref="Word.Field"/>s.</param>
        /// <returns>All <see cref="Word.Field"/>s in the specified <see cref="Word.Selection"/>.</returns>
        /// <remarks>
        /// The code has been taken from
        /// <a href="http://stackoverflow.com/questions/11243752/check-what-field-was-clicked-in-ms-word">this</a>
        /// Stack Overflow question.
        /// </remarks>
        public static IEnumerable<Word.Field> AllFields(this Word.Selection selection)
        {
            // TODO Improve performance.
            foreach (Word.Field field in selection.Document.Fields)
            {
                string fieldResult = field.Result.Text;
                
                if (null == fieldResult)
                {
                    fieldResult = string.Empty;
                }

                int fieldStart = field.Code.FormattedText.Start;
                int fieldEnd = field.Code.FormattedText.End + fieldResult.Length;

                if (!((fieldStart < selection.Start) & (fieldEnd < selection.Start)
                    | (fieldStart > selection.End) & (fieldEnd > selection.End)))
                {
                    yield return field;
                }
            }
        }

        /// <summary>
        /// Retrieves all <i>INCLUDE[...]</i> <see cref="Word.Field"/>s in the specified <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="selection">The <see cref="Word.Selection"/> to check.</param>
        /// <returns>
        /// All <i>INCLUDE[...]</i> <see cref="Word.Field"/>s in the specified <see cref="Word.Selection"/>.
        /// </returns>
        public static IEnumerable<Word.Field> AllIncludeFields(this Word.Selection selection)
        {
            // TODO Improve performance.
            return from f in AllFields(selection)
                   where f.IsTypeInclude()
                   select f;
        }

        /// <summary>
        /// Retrieves all <i>IncludeText</i> <see cref="Word.Field"/>s in the specified <see cref="Word.Selection"/>.
        /// </summary>
        /// <param name="selection">The <see cref="Word.Selection"/> to check.</param>
        /// <returns>
        /// All <i>IIncludeText</i> <see cref="Word.Field"/>s in the specified <see cref="Word.Selection"/>.
        /// </returns>
        public static IEnumerable<Word.Field> AllIncludeTextFields(this Word.Selection selection)
        {
            // TODO Improve performance.
            return from f in AllFields(selection)
                   where f.CanUpdateSource()
                   select f;
        }
    }
}
