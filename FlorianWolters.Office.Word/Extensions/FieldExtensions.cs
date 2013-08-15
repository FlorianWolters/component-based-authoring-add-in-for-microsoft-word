//------------------------------------------------------------------------------
// <copyright file="FieldExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Extensions
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The static class <see cref="FieldExtensions"/> contains extension methods for a Microsoft Word field,
    /// represented by an object of the class <see cref="Word.Field"/>.
    /// </summary>
    public static class FieldExtensions
    {
        /// <summary>
        /// Determines whether the specified <see cref="Word.Field"/> can be updated.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to check.</param>
        /// <returns><c>true</c> if the specified <see cref="Word.Field"/> can be updated; <c>false</c> otherwise.
        /// </returns>
        public static bool CanUpdate(this Word.Field field)
        {
            return field.Locked = false;
        }

        /// <summary>
        /// Determines whether the specified <see cref="Word.Field"/> is an <i>INCLUDE</i>, <i>INCLUDETEXT</i> or
        /// <i>INCLUDEPICTURE</i> field.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to check.</param>
        /// <returns>
        /// <c>true</c> if the specified <see cref="Word.Field"/> is an <i>INCLUDE</i>, <i>INCLUDETEXT</i> or
        /// <i>INCLUDEPICTURE</i> field; <c>false</c> otherwise.
        /// </returns>
        public static bool IsTypeInclude(this Word.Field field)
        {
            return field.Type == Word.WdFieldType.wdFieldIncludeText
                || field.Type == Word.WdFieldType.wdFieldInclude
                || field.Type == Word.WdFieldType.wdFieldIncludePicture;
        }

        /// <summary>
        /// Determines whether the source of the specified <see cref="Word.Field"/> can be updated.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to check.</param>
        /// <returns>
        /// <c>true</c> if the source of the specified <see cref="Word.Field"/> can be updated; <c>false</c> otherwise.
        /// </returns>
        public static bool CanUpdateSource(this Word.Field field)
        {
            return field.Type == Word.WdFieldType.wdFieldIncludeText
                || field.Type == Word.WdFieldType.wdFieldInclude;
        }
    }
}
