//------------------------------------------------------------------------------
// <copyright file="XmlDocumentListBoxItem.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.Models
{
    using System.Xml;

    /// <summary>
    /// The class <see cref="XmlDocumentListBoxItem"/> allows to use <see cref="XmlDocument"/> objects inside a <see
    /// cref="System.Windows.Forms.ListBox"/>.
    /// <para>
    /// The default namespace of the <see cref="XmlDocument"/> is used as the text if this object is passed to a <see
    /// cref="System.Windows.Forms.ListBox"/> object.
    /// </para>
    /// </summary>
    internal class XmlDocumentListBoxItem
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XmlDocumentListBoxItem"/> class with the specified <see
        /// cref="XmlDocument"/>.
        /// </summary>
        /// <param name="xmlDocument">The <see cref="XmlDocument"/> to wrap.</param>
        public XmlDocumentListBoxItem(XmlDocument xmlDocument)
        {
            this.XmlDocument = xmlDocument;
        }

        /// <summary>
        /// Gets the wrapped <see cref="XmlDocument"/>.
        /// </summary>
        public XmlDocument XmlDocument { get; private set; }

        /// <summary>
        /// Returns a string representation of this <see cref="XmlDocumentListBoxItem"/>.
        /// </summary>
        /// <remarks>
        /// The return value of this method is used as the text if this object is passed to a <see
        /// cref="System.Windows.Forms.ListBox"/> object.
        /// </remarks>
        /// <returns>The string representation.</returns>
        public override string ToString()
        {
            string defaultNamespace = this.XmlDocument.DocumentElement.NamespaceURI;

            if (string.Empty == defaultNamespace)
            {
                defaultNamespace = "(no default namespace)";
            }

            return defaultNamespace;
        }
    }
}
