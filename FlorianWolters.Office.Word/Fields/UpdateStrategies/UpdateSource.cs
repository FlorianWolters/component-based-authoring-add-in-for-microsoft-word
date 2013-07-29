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

    public class UpdateSource : IUpdateStrategy
    {
        private const int ErrorCodeReadOnly = -2146823133;

        public void Update(Word.Field field)
        {
            if (field.CanUpdateSource())
            {
                try
                {
                    field.UpdateSource();
                }
                catch (COMException ex)
                {
                    if (ErrorCodeReadOnly == ex.ErrorCode)
                    {
                        throw new ReadOnlyDocumentException(
                            "The included document is read-only.");
                    }
                }
            }
            else
            {
                throw new ArgumentException("Invalid field type.");
            }
        }
    }
}
