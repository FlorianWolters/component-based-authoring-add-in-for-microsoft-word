//------------------------------------------------------------------------------
// <copyright file="CustomXMLNodeMappingStrategy.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.ContentControls.MappingStrategies
{
    using System;
    using Office = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The abstract class <see cref="CustomXMLNodeMappingStrategy"/> allows to map a <see cref="Office.CustomXMLNode"/>
    /// to one or more <see cref="Word.ContentControl"/>s.
    /// </summary>
    public abstract class CustomXMLNodeMappingStrategy : IMappingStrategy
    {
        /// <summary>
        /// The root <see cref="Office.CustomXMLNode"/> which determines the data to map.
        /// </summary>
        protected readonly Office.CustomXMLNode RootNode;

        /// <summary>
        /// Used to create instances of <see cref="Word.ContentControl"/>.
        /// </summary>
        protected readonly ContentControlFactory ContentControlFactory;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomXMLNodeMappingStrategy"/> class.
        /// </summary>
        /// <param name="rootNode">The root <see cref="Office.CustomXMLNode"/> which determines the data to map.</param>
        /// <param name="contentControlFactory">Used to create instances of <see cref="Word.ContentControl"/>.</param>
        /// <exception cref="ArgumentNullException">If <c>rootNode</c> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">If <c>contentControlFactory</c> is <c>null</c>.</exception>
        protected CustomXMLNodeMappingStrategy(
            Office.CustomXMLNode rootNode,
            ContentControlFactory contentControlFactory)
        {
            if (null == rootNode)
            {
                throw new ArgumentNullException("rootNode");
            }

            if (null == contentControlFactory)
            {
                throw new ArgumentNullException("contentControlFactory");
            }

            this.RootNode = rootNode;
            this.ContentControlFactory = contentControlFactory;
            this.ThrowContentControlMappingExceptionIfNotMappable();
        }

        /// <summary>
        /// Maps the data of the <i>Strategy</i> to <see cref="Word.ContentControl"/>s which are created in the
        /// specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="range">The <see cref="Word.Range"/> to use.</param>
        /// <returns>The <see cref="Word.Range"/> which has been created.</returns>
        /// <exception cref="ArgumentNullException">If <c>range</c> is <c>null</c>.</exception>
        public abstract Word.Range MapToCustomControlsIn(Word.Range range);

        /// <summary>
        /// Determines whether the root <see cref="Office.CustomXMLNode"/> can be mapped with this <i>Strategy</i> or
        /// not.
        /// </summary>
        /// <returns><c>true</c> if the root node can be mapped; <c>false</c> otherwise.</returns>
        protected abstract bool IsRootNodeMappable();

        /// <summary>
        /// Creates and inserts a new <see cref="Word.ContentControl"/> which is mapped to the specified <see
        /// cref="Office.CustomXMLNode"/> in the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to map the <see cref="Word.ContentControl"/> to.</param>
        /// <param name="range">The <see cref="Word.Range"/> to insert the <see cref="Word.ContentControl"/> in.</param>
        /// <returns>The <see cref="Word.Range"/> of the newly created <see cref="Word.ContentControl"/>.</returns>
        protected Word.ContentControl CreateContentControl(
            Office.CustomXMLNode node,
            Word.Range range)
        {
            Word.WdContentControlType contentControlType = Word.WdContentControlType.wdContentControlText;
            bool isBoolean = false;

            if (bool.TryParse(node.NodeValue, out isBoolean))
            {
                // If the value of the CustomXMLNode is either "true" or "false" a checkbox content control can be
                // created instead of a text content control.
                contentControlType = Word.WdContentControlType.wdContentControlCheckBox;
            }

            // Don't lock the content control, since otherwise the user cannot move the content control in the document.
            return this.ContentControlFactory.CreateContentControl(
                contentControlType,
                node,
                range,
                lockContents: true);
        }

        /// <summary>
        /// Throws a <see cref="ContentControlMappingException"/> if the root <see cref="Office.CustomXMLNode"/> cannot
        /// be mapped with this <i>Strategy</i>.
        /// </summary>
        /// <exception cref="ContentControlMappingException">If the root node cannot be mapped.</exception>
        private void ThrowContentControlMappingExceptionIfNotMappable()
        {
            if (!this.IsRootNodeMappable())
            {
                throw new ContentControlMappingException(
                    "The specified CustomXMLNode cannot be mapped with this strategy.");
            }
        }
    }
}
