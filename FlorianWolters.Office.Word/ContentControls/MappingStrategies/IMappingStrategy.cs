//------------------------------------------------------------------------------
// <copyright file="IMappingStrategy.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.ContentControls.MappingStrategies
{
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The interface <see cref="IMappingStrategy"/> specified that the implementing class is able to map data to one or
    /// more <see cref="Word.ContentControl"/>s in a Microsoft Word document.
    /// </summary>
    public interface IMappingStrategy
    {
        /// <summary>
        /// Maps the data of the <i>Strategy</i> to <see cref="Word.ContentControl"/>s which are created in the
        /// specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> to use.</param>
        /// <returns>The <see cref="Word.Range"/> which has been created.</returns>
        /// <exception cref="System.ArgumentNullException">If <c>range</c> is <c>null</c>.</exception>
        Word.Range MapToCustomControlsIn(Word.Range range); 
    }
}
