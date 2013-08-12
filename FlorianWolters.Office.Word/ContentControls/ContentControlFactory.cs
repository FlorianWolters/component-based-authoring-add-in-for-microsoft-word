//------------------------------------------------------------------------------
// <copyright file="ContentControlFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.ContentControls
{
    using System;
    using System.Runtime.InteropServices;
    using Office = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="ContentControlFactory"/> allows to create new instances of the class <see
    /// cref="ContentControl"/>.
    /// </summary>
    public class ContentControlFactory
    {
        /// <summary>
        /// The Microsoft Word document in which to create the content controls.
        /// </summary>
        private readonly Word.Document document;

        /// <summary>
        /// Initializes a new instance of the <see cref="ContentControlFactory"/> class.
        /// </summary>
        /// <param name="document">The Microsoft Word document in which to create the content controls.</param>
        /// <exception cref="ArgumentNullException">If <c>document</c> is <c>null</c>.</exception>
        public ContentControlFactory(Word.Document document)
        {
            if (null == document)
            {
                throw new ArgumentNullException("document");
            }

            this.document = document;
        }

        /// <summary>
        /// Creates a new instance of <see cref="ContentControl"/>.
        /// </summary>
        /// <param name="contentControlType">The type of content control to create.</param>
        /// <param name="customXMLNode">The custom XML node to map to the content control.</param>
        /// <param name="range">The range where to insert the content control.</param>
        /// <param name="lockContents"><c>true</c> if the contents of the content control should be locked.</param>
        /// <param name="lockControl"><c>true</c> if the content control should be locked.</param>
        /// <returns>A newly created <see cref="ContentControl"/>.</returns>
        /// <exception cref="ContentControlCreationException">If a content control cannot be created in the specified range.</exception>
        public Word.ContentControl CreateContentControl(
            Word.WdContentControlType contentControlType,
            Office.CustomXMLNode customXMLNode = null,
            Word.Range range = null,
            bool lockContents = true,
            bool lockControl = false)
        {
            Word.ContentControl result = null;

            try
            {
                result = this.document.ContentControls.Add(contentControlType, range);

                result.Tag = result.ID;
                this.MapContentControlToCustomXMLNode(result, customXMLNode);
                result.LockContents = lockContents;
                result.LockContentControl = lockControl;
            }
            catch (COMException ex)
            {
                throw new ContentControlCreationException(
                    "Unable to create a content control in the specified range.",
                    ex);
            }

            return result;
        }

        /// <summary>
        /// Maps the specified <see cref="Word.ContentControl"/> to the specified <see cref="Office.CustomXMLNode"/>
        /// </summary>
        /// <param name="contentControl">The content control to map.</param>
        /// <param name="customXMLNode">The custom XML node to map.</param>
        /// <exception cref="ContentControlMappingException">If unable to map the content control to the custom XML node.</exception>
        private void MapContentControlToCustomXMLNode(
            Word.ContentControl contentControl,
            Office.CustomXMLNode customXMLNode)
        {
            if (null != customXMLNode && Word.WdContentControlType.wdContentControlText == contentControl.Type)
            {
                if (!contentControl.XMLMapping.SetMappingByNode(customXMLNode))
                {
                    throw new ContentControlMappingException(
                        "Unable to map the content control to the custom XML part.");
                }
            }
        }
    }
}
