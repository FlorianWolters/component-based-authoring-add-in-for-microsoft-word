//------------------------------------------------------------------------------
// <copyright file="UpdateSource.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using System;
    using System.Runtime.InteropServices;
    using FlorianWolters.Office.Word.Extensions;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="UpdateSource"/> saves the changes made to the results of an INCLUDETEXT (or INCLUDE) <see
    /// cref="Word.Field"/> back to the source document.
    /// </summary>
    public class UpdateSource : IUpdateStrategy
    {
        /// <summary>
        /// The error code which determines that the source document is read-only.
        /// </summary>
        private const int ErrorCodeReadOnly = -2146823133;

        /// <summary>
        /// Updates the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to update.</param>
        public void Update(Word.Field field)
        {
            if (!field.CanUpdateSource())
            {
                throw new ArgumentException("Invalid field type.");
            }

            try
            {
                field.UpdateSource();
            }
            catch (COMException ex)
            {
                if (ErrorCodeReadOnly == ex.ErrorCode)
                {
                    throw new ReadOnlyDocumentException("The included document is read-only.");
                }
            }
        }
    }
}
