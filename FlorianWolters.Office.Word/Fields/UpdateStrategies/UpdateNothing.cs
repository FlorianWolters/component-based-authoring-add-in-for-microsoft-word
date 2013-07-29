//------------------------------------------------------------------------------
// <copyright file="UpdateNothing.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using Word = Microsoft.Office.Interop.Word;

    public class UpdateNothing : IUpdateStrategy
    {
        public void Update(Word.Field field)
        {
            // NOOP
        }
    }
}
