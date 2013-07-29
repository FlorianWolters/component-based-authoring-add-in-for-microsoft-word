//------------------------------------------------------------------------------
// <copyright file="UpdateTarget.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using System;
    using FlorianWolters.Office.Word.Extensions;
    using Word = Microsoft.Office.Interop.Word;

    public class UpdateTarget : IUpdateStrategy
    {
        public void Update(Word.Field field)
        {
            if (field.CanUpdate())
            {
                field.Update();
            }
            else
            {
                throw new ArgumentException("Invalid field type.");
            }
        }
    }
}
