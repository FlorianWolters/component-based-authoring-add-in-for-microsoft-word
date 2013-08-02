//------------------------------------------------------------------------------
// <copyright file="IUpdateStrategy.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using Word = Microsoft.Office.Interop.Word;

    public interface IUpdateStrategy
    {
        /// <summary>
        /// Updates the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to update.</param>
        void Update(Word.Field field);
    }
}
