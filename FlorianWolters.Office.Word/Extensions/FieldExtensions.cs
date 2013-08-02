//------------------------------------------------------------------------------
// <copyright file="FieldExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Extensions
{
    using Word = Microsoft.Office.Interop.Word;

    public static class FieldExtensions
    {
        public static bool CanUpdate(this Word.Field field)
        {
            return field.Locked = false;
        }

        public static bool IsTypeInclude(this Word.Field field)
        {
            return field.Type == Word.WdFieldType.wdFieldIncludeText
                || field.Type == Word.WdFieldType.wdFieldInclude
                || field.Type == Word.WdFieldType.wdFieldIncludePicture;
        }

        public static bool CanUpdateSource(this Word.Field field)
        {
            return field.Type == Word.WdFieldType.wdFieldIncludeText
                || field.Type == Word.WdFieldType.wdFieldInclude;
        }
    }
}
