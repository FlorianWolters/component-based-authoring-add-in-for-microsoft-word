//------------------------------------------------------------------------------
// <copyright file="SelectedXmlDocumentChangedEventArgs.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.EventArgs
{
    using System;
    using System.Xml;

    /// <summary>
    /// The class <see cref="SelectedXmlDocumentChangedEventArgs"/> provides a <see cref="XmlDocument"/> object as data
    /// for an event.
    /// </summary>
    public class SelectedXmlDocumentChangedEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SelectedXmlDocumentChangedEventArgs"/> class.
        /// </summary>
        /// <param name="xmlDocument">The <see cref="XmlDocument"/>.</param>
        public SelectedXmlDocumentChangedEventArgs(XmlDocument xmlDocument)
        {
            this.XmlDocument = xmlDocument;
        }

        /// <summary>
        /// Gets the <see cref="XmlDocument"/>.
        /// </summary>
        public XmlDocument XmlDocument { get; private set; }
    }
}
