﻿//------------------------------------------------------------------------------
// <copyright file="UpdateTarget.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="UpdateTarget"/> updates the result of a <see cref="Word.Field"/>.
    /// </summary>
    public class UpdateTarget : IUpdateStrategy
    {
        /// <summary>
        /// Updates the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to update.</param>
        public void Update(Word.Field field)
        {
            field.Update();
        }
    }
}
