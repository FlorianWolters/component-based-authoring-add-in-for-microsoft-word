//------------------------------------------------------------------------------
// <copyright file="XMLDataUserControl.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.UserControls
{
    using System.Windows.Forms;
    using System.Xml;
    using FlorianWolters.Windows.Forms.XML.Extensions;

    /// <summary>
    /// The class <see cref="XMLDataUserControl"/> implements an <see cref="UserControl"/> which displays the data of a
    /// <see cref="XmlNode"/>.
    /// </summary>
    public partial class XMLDataUserControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XMLDataUserControl"/> class.
        /// </summary>
        public XMLDataUserControl()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Handles the event that occurs after a <see cref="TreeNode"/> is selected.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        internal void OnAfterSelectTreeViewStructure(object sender, TreeViewEventArgs e)
        {
            XmlNode xmlNode = (XmlNode)e.Node.Tag;

            this.listViewData.Items.Clear();

            ListViewItem listViewItem = new ListViewItem(xmlNode.LocalName);

            // TODO Apply "Replace Conditional with Polymorphism".
            switch (xmlNode.NodeType)
            {
                case XmlNodeType.Element:
                    if (xmlNode.IsLeaf())
                    {
                        listViewItem.SubItems.Add(xmlNode.InnerText);
                    }
                    else
                    {
                        listViewItem.SubItems.Add(string.Empty);
                    }

                    break;
                case XmlNodeType.Attribute:
                    listViewItem.SubItems.Add(xmlNode.InnerText);
                    break;
                default:
                    listViewItem.SubItems.Add(string.Empty);
                    break;
            }

            listViewItem.SubItems.Add(xmlNode.NodeType.ToString());

            this.listViewData.Items.Add(listViewItem);
        }
    }
}
