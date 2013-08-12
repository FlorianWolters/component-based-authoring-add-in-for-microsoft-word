//------------------------------------------------------------------------------
// <copyright file="XMLBrowserForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.Forms
{
    using System.Windows.Forms;
    using System.Xml;
    using FlorianWolters.Windows.Forms.XML.EventArgs;

    /// <summary>
    /// The class <see cref="XMLBrowserForm"/> implements a Windows Form which allows to select one <see
    /// cref="XmlNode"/> from multiple <see cref="XmlDocument"/>s.
    /// </summary>
    public partial class XMLBrowserForm : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XMLBrowserForm"/> class.
        /// </summary>
        public XMLBrowserForm()
        {
            this.InitializeComponent();

            // TODO This form is responsible to register the correct event handlers, so that the user controls do not
            // depend on each other. I am unsure if this is the "correct" approach.
            this.xmlNamespaceUserControl.SelectedXmlDocumentChanged += this.xmlStructureUserControl.OnXmlDocumentSelected;
            this.xmlStructureUserControl.treeViewStructure.AfterSelect += this.xmlDataUserControl.OnAfterSelectTreeViewStructure;
            this.xmlStructureUserControl.SelectedTreeNodeChanged += this.xPathUserControl.OnAfterSelectTreeView;
        }

        // TODO This is ugly, since it exposes the XPathEventArgs class.

        /// <summary>
        /// Gets the XPath result for the selected XML node.
        /// </summary>
        public XPathEventArgs ResultXPath
        {
            get
            {
                return this.xmlStructureUserControl.XPathEventArgs;
            }
        }

        /// <summary>
        /// Gets the selected <see cref="XmlDocument"/>.
        /// </summary>
        public XmlDocument ResultXmlDocument
        {
            get
            {
                return this.xmlNamespaceUserControl.SelectedXmlDocument;
            }
        }

        /// <summary>
        /// Adds the specified array of <see cref="XmlDocument"/>s to this <see cref="XMLBrowserForm"/>.
        /// </summary>
        /// <param name="xmlDocuments">The array of <see cref="XmlDocument"/> to add.</param>
        public void AddXmlDocuments(XmlDocument[] xmlDocuments)
        {
            foreach (XmlDocument xmlDocument in xmlDocuments)
            {
                this.xmlNamespaceUserControl.AddXmlDocument(xmlDocument);
            }
        }

        /// <summary>
        /// Adds the specified <see cref="XmlDocument"/> to this <see cref="XMLBrowserForm"/>.
        /// </summary>
        /// <param name="xmlDocument">The <see cref="XmlDocument"/> to add.</param>
        public void AddXmlDocument(XmlDocument xmlDocument)
        {
            this.xmlNamespaceUserControl.AddXmlDocument(xmlDocument);
        }
    }
}
