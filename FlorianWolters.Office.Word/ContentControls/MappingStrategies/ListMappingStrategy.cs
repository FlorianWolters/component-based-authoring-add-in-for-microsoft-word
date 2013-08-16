//------------------------------------------------------------------------------
// <copyright file="ListMappingStrategy.cs" company="Florian Wolters">
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
    /// The class <see cref="ListMappingStrategy"/> maps <see
    /// cref="Office.CustomXMLNode"/>s to a list of <see cref="Office.ContentControl"/>s in a Microsoft Word document.
    /// <para>
    /// This class only outputs XML nodes of type element or attribute.
    /// Each node name follows a content control with the value of the node.
    /// * Example for an element: <c>&lt;name&gt; : value</c>.
    /// * Example for an attribute: <c>name = value</c>.
    /// </para>
    /// </summary>
    public class ListMappingStrategy : CustomXMLNodeMappingStrategy
    {
        /// <summary>
        /// The number of list templates to use.
        /// <para>
        /// If the current level of the list exceeds this number, the first list
        /// template is used again.
        /// </para>
        /// </summary>
        private const int NumberOfListTemplatesToUse = 3;

        /// <summary>
        /// The gallery of the list format to use.
        /// </summary>
        private readonly Word.ListGallery listGallery;

        /// <summary>
        /// The level of the root <see cref="Office.CustomXMLNode"/>.
        /// </summary>
        private readonly int rootNodeLevel = 0;

        /// <summary>
        /// The level of the current <see cref="Office.CustomXMLNode"/>.
        /// </summary>
        private int currentLevel;

        /// <summary>
        /// Initializes a new instance of the <see cref="ListMappingStrategy"/> class.
        /// </summary>
        /// <param name="rootNode">The root <see cref="Office.CustomXMLNode"/> which determines the data to map.</param>
        /// <param name="contentControlFactory">Used to create instances of <see cref="Office.ContentControl"/>.</param>
        /// <param name="listGallery">The gallery of the list format used to style the output.</param>
        /// <exception cref="ArgumentNullException">If <c>rootNode</c> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">If <c>contentControlFactory</c> is <c>null</c>.</exception>
        /// <exception cref="ArgumentNullException">If <c>listGallery</c> is <c>null</c>.</exception>
        public ListMappingStrategy(
            Office.CustomXMLNode rootNode,
            ContentControlFactory contentControlFactory,
            Word.ListGallery listGallery)
            : base(rootNode, contentControlFactory)
        {
            if (null == listGallery)
            {
                throw new ArgumentNullException("listGallery");
            }

            this.listGallery = listGallery;
            this.rootNodeLevel = rootNode.GetLevel();
        }

        /// <summary>
        /// Maps the data of the <i>Strategy</i> to <see cref="Office.ContentControl"/>s which are created in the
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

            // TODO Wow, this is serious crap code.
            Word.Range oldRange = range;

            range = range.Application.ActiveDocument.Content;

            // TODO The handling of the Word.Range within this class is incorrect. Therfore this method only returns the
            // correct result if Document.Content.Range is passed as the argument. The returned Word.Range is correct.
            int rangeStart = range.End;
            range.InsertParagraphAfter();

            // Disable screen updating of the Microsoft Word application to improve the performance.
            range.Application.ScreenUpdating = false;
            this.MapCustomXMLNodeToContentControls(this.RootNode, range);
            range.Application.ScreenUpdating = true;

            Word.Range insertedRange = range.Application.ActiveDocument.Range(rangeStart - 1, range.End);
            insertedRange.Select();
            insertedRange.Application.Selection.Copy();
            insertedRange.Delete();
            oldRange.Paste();

            return oldRange;
        }

        /// <summary>
        /// Determines whether the root <see cref="Office.CustomXMLNode"/> can be mapped with this <i>Strategy</i> or
        /// not.
        /// </summary>
        /// <returns><c>true</c> if the root node can be mapped; <c>false</c> otherwise.</returns>
        protected override bool IsRootNodeMappable()
        {
            // This strategy can map every CustomXMLNode structure.
            return true;
        }

        /// <summary>
        /// Maps the specified <see cref="Office.CustomXMLNode"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to map.</param>
        /// <param name="range">The <see cref="Word.Range"/> to use.</param>
        private void MapCustomXMLNodeToContentControls(Office.CustomXMLNode node, Word.Range range)
        {
            // Counting the level of each node isn't efficient, but easier to handle in contrast to manage the level
            // manually.
            this.currentLevel = node.GetLevel() - (this.rootNodeLevel - 1);

            switch (node.NodeType)
            {
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement:
                    if (!node.HasChildNodes() && 0 == node.Attributes.Count)
                    {
                        // Ignore empty element nodes without attributes, e.g. "<empty/>".
                        break;
                    }

                    if (node.HasChildNodes() && node.FirstChild.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeText)
                    {
                        // If the CustomXMLNode is of type msoCustomXMLNodeElement and its first child node is a text
                        // node, map the CustomXMLNode to a new content control.
                        this.InsertLabelAndContentControlParagraph(node, range);
                    }
                    else
                    {
                        // The CustomXMLNode is either an empty element node or does contain one or more element nodes. 
                        this.InsertLabelParagraph(node, range);
                    }

                    if (node.Attributes.Count > 0)
                    {
                        this.MapCustomXMLNodesToContentControls(node.Attributes, range);
                    }

                    // A CustomXMLNode of type msoCustomXMLNodeElement always has at least one child CustomXMLNode.
                    this.MapCustomXMLNodesToContentControls(node.ChildNodes, range);
                    break;
                case Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute:
                    // If the CustomXMLNode is of type msoCustomXMLNodeAttribute, map it to a new content control.
                    this.InsertLabelAndContentControlParagraph(node, range);
                    break;
                default:
                    // NOOP
                    // We ignore the other CustomXMLNode types.
                    break;
            }
        }

        /// <summary>
        /// Maps the specified <see cref="Office.CustomXMLNodes"/> to the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="nodes">The <see cref="Office.CustomXMLNodes"/> to map.</param>
        /// <param name="range">The <see cref="Word.Range"/> to use.</param>
        private void MapCustomXMLNodesToContentControls(Office.CustomXMLNodes nodes, Word.Range range)
        {
            foreach (Office.CustomXMLNode node in nodes)
            {
                this.MapCustomXMLNodeToContentControls(node, range);
            }
        }

        /// <summary>
        /// Inserts a <see cref="Word.Paragraph"/> which contains the base name from the specified <see
        /// cref="Office.CustomXMLNode"/> in the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNodes"/> whose base name to use.</param>
        /// <param name="range">The <see cref="Word.Range"/> to use.</param>
        /// <returns>The <see cref="Word.Range"/> of the newly created <see cref="Word.Paragraph."/></returns>
        private Word.Range InsertLabelParagraph(Office.CustomXMLNode node, Word.Range range)
        {
            Word.Paragraph paragraph = range.Paragraphs.Add();
            Word.Range contentControlRange = paragraph.Range;

            // The list-formatting characteristics must be retrieved, before we set the text of the range. Otherwise the
            // list is not indented.
            Word.ListFormat listFormat = paragraph.Range.ListFormat;

            string text = string.Empty;

            if (node.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement)
            {
                text = "<" + node.BaseName + "> ";
            }
            else if (node.NodeType == Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute)
            {
                text = node.BaseName + " = ";
            }

            // The text of the range must be set, before the list-formatting characteristic are applied.
            paragraph.Range.Text = text;

            this.ApplyListTemplate(listFormat);

            contentControlRange.Move(Unit: Word.WdUnits.wdParagraph);
            contentControlRange.InsertParagraphAfter();

            return contentControlRange;
        }

        /// <summary>
        /// Inserts a new <see cref="Word.Paragraph"/> which contains the base name from the specified <see
        /// cref="Office.CustomXMLNode"/> and inserts a new <see cref="Word.ContentControl"/> which is mapped to the
        /// specified <see cref="Office.CustomXMLNode"/> in the specified <see cref="Word.Range"/>.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> whose data to use.</param>
        /// <param name="range">The <see cref="Word.Range"/> to insert both elements in.</param>
        /// <returns>The <see cref="Word.Range"/> of the newly created <see cref="Word.ContentControl."/></returns>
        private Word.Range InsertLabelAndContentControlParagraph(Office.CustomXMLNode node, Word.Range range)
        {
            Word.Range paragraphRange = this.InsertLabelParagraph(node, range);
            this.CreateContentControl(node, paragraphRange);

            return paragraphRange;
        }

        /// <summary>
        /// Applies the specified <see cref="Word.ListFormat"/> to the current selection.
        /// </summary>
        /// <param name="listFormat">The <see cref="Word.ListFormat"/> to apply.</param>
        private void ApplyListTemplate(Word.ListFormat listFormat)
        {
            listFormat.ApplyListTemplateWithLevel(
                this.GetListTemplateFromListGallery(),
                ContinuePreviousList: true,
                ApplyTo: Word.WdListApplyTo.wdListApplyToSelection,
                DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel: this.currentLevel);
        }

        /// <summary>
        /// Returns a <see cref="Word.ListTemplate"/> from the <see cref="Word.Gallery"/> which corresponds to the
        /// current level of the list.
        /// </summary>
        /// <returns>A <see cref="Word.ListTemplate"/>.</returns>
        private Word.ListTemplate GetListTemplateFromListGallery()
        {
            // We only switch between a constant number of list templates to avoid a COMException if an invalid index
            // (greater than the maximum number of available list templates) is specified.
            int listLevel = this.currentLevel;
            int newListLevel = this.currentLevel % NumberOfListTemplatesToUse;

            if (this.currentLevel > NumberOfListTemplatesToUse && 0 != newListLevel)
            {
                listLevel = newListLevel;
            }

            return this.listGallery.ListTemplates[listLevel];
        }
    }
}
