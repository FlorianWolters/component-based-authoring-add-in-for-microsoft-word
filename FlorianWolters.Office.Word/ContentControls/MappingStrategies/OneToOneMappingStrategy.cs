//------------------------------------------------------------------------------
// <copyright file="OneToOneMappingStrategy.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.ContentControls.MappingStrategies
{
    using System;
    using FlorianWolters.Office.Word.Extensions;
    using Office = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="OneToOneMappingStrategy"/> maps one <see cref="Office.CustomXMLNode"/> to one <see
    /// cref="Word.ContentControl"/> in a Microsoft Word document.
    /// </summary>
    public class OneToOneMappingStrategy : CustomXMLNodeMappingStrategy
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OneToOneMappingStrategy"/> class.
        /// </summary>
        /// <param name="rootNode">The root <see cref="Office.CustomXMLNode"/> which determines the data to map.</param>
        /// <param name="contentControlFactory">Used to create instances of <see cref="Word.ContentControl"/>.</param>
        /// <exception cref="ArgumentNullException">If <c>rootNode</c> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">If <c>contentControlFactory</c> is <c>null</c>.</exception>
        public OneToOneMappingStrategy(Office.CustomXMLNode rootNode, ContentControlFactory contentControlFactory)
            : base(rootNode, contentControlFactory)
        {
        }

        /// <summary>
        /// Maps the data of the <i>Strategy</i> to <see cref="Word.ContentControl"/>s which are created in the
        /// specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> to use.</param>
        /// <returns>The <see cref="Word.Range"/> which has been created.</returns>
        /// <exception cref="ArgumentNullException">If <c>range</c> is <c>null</c>.</exception>
        public override Word.Range MapToCustomControlsIn(Word.Range range)
        {
            if (null == range)
            {
                throw new ArgumentNullException("range");
            }

            return this.CreateContentControl(this.RootNode, range).Range;
        }

        /// <summary>
        /// Determines whether the root <see cref="Office.CustomXMLNode"/> can be mapped with this <i>Strategy</i> or
        /// not.
        /// </summary>
        /// <returns><c>true</c> if the root node can be mapped; <c>false</c> otherwise.</returns>
        protected override bool IsRootNodeMappable()
        {
            return this.RootNode.IsAttribute() || this.RootNode.IsElement();
        }
    }
}
