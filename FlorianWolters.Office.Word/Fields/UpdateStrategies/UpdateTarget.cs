//------------------------------------------------------------------------------
// <copyright file="UpdateTarget.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using Word = Microsoft.Office.Interop.Word;

    public class UpdateTarget : IUpdateStrategy
    {
        /// <summary>
        /// Updates the specified <see cref="Word.Field"/>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to update.</param>
        public void Update(Word.Field field)
        {
            // If a field is updated if its field codes are visible, the field
            // codes are replaced by the field result. Therefore we do have to
            // manually keep track of the UI state.
            bool showCodes = field.ShowCodes;
            field.Update();
            field.ShowCodes = showCodes;
        }
    }
}
