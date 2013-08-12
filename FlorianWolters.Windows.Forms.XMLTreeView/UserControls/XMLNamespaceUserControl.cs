//------------------------------------------------------------------------------
// <copyright file="XMLNamespaceUserControl.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.UserControls
{
    using System;
    using System.Windows.Forms;
    using System.Xml;
    using FlorianWolters.Windows.Forms.XML.EventArgs;
    using FlorianWolters.Windows.Forms.XML.Models;

    /// <summary>
    /// The class <see cref="XMLNamespaceUserControl"/> implements a <see cref="UserControl"/> which allows to select a
    /// <see cref="XmlDocument"/> instance in a <see cref="ListBox"/>.
    /// </summary>
    public partial class XMLNamespaceUserControl : UserControl
    {
        /// <summary>
        /// The index of the last selected item from the <see cref="ListBox"/>.
        /// </summary>
        private int previousSelectedIndex = -1;

        /// <summary>
        /// Initializes a new instance of the <see cref="XMLNamespaceUserControl"/> class.
        /// </summary>
        public XMLNamespaceUserControl()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Handles the event that occurs after the selected <see cref="XmlDocument"/> has changed.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event data.</param>
        public delegate void SelectedXmlDocumentChangedHandler(
            object sender,
            SelectedXmlDocumentChangedEventArgs e);

        /// <summary>
        /// Raised after the selected <see cref="XmlDocument"/> has changed.
        /// </summary>
        public event EventHandler<SelectedXmlDocumentChangedEventArgs> SelectedXmlDocumentChanged;

        /// <summary>
        /// Gets the currently selected <see cref="XmlDocument"/>.
        /// </summary>
        public XmlDocument SelectedXmlDocument
        {
            get
            {
                return ((XmlDocumentListBoxItem)this.listBoxNamespaces.SelectedItem).XmlDocument;
            }
        }

        /// <summary>
        /// Adds the specified <see cref="XmlDocument"/> to this <see cref="XMLNamespaceUserControl"/>.
        /// </summary>
        /// <param name="xmlDocument">The <see cref="XmlDocument"/> do add.</param>
        public void AddXmlDocument(XmlDocument xmlDocument)
        {
            XmlDocumentListBoxItem listBoxItem = new XmlDocumentListBoxItem(xmlDocument);

            this.listBoxNamespaces.Items.Add(listBoxItem);
        }

        /// <summary>
        /// Handles the event that occurs when the <see cref="SelectedIndex"/> property or the <see
        /// cref="SelectedIndices"/> collection of the <see cref="ListBox"/> has changed.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnSelectedIndexChangedListBox(object sender, System.EventArgs e)
        {
            if (this.previousSelectedIndex != this.listBoxNamespaces.SelectedIndex)
            {
                XmlDocumentListBoxItem selectedItem = (XmlDocumentListBoxItem)this.listBoxNamespaces.SelectedItem;

                if (null != selectedItem)
                {
                    this.RaiseSelectedXmlDocumentChangedEvent(selectedItem.XmlDocument);
                }

                this.previousSelectedIndex = this.listBoxNamespaces.SelectedIndex;
            }
        }

        /// <summary>
        /// Raises an event which signals that the selected <see cref="XmlDocument"/> has changed.
        /// </summary>
        /// <param name="xmlDocument">The currently selected <see cref="XmlDocument"/>.</param>
        private void RaiseSelectedXmlDocumentChangedEvent(XmlDocument xmlDocument)
        {
            if (null != this.SelectedXmlDocumentChanged)
            {
                // Create the event arguments.
                SelectedXmlDocumentChangedEventArgs args = new SelectedXmlDocumentChangedEventArgs(xmlDocument);

                // Raise the event.
                this.SelectedXmlDocumentChanged(this, args);
            }
        }
    }
}
