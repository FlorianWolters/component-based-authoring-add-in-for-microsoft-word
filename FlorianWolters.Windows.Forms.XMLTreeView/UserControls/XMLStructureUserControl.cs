//------------------------------------------------------------------------------
// <copyright file="XMLStructureUserControl.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.UserControls
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Text;
    using System.Windows.Forms;
    using System.Xml;
    using FlorianWolters.Windows.Forms.XML.EventArgs;
    using FlorianWolters.Windows.Forms.XML.Extensions;
    using FlorianWolters.Windows.Forms.XML.Models;

    /// <summary>
    /// The class <see cref="XMLStructureUserControl"/> implements a <see cref="UserControl"/> which allows to browse a
    /// XML structure represented via a <see cref="XmlNode"/> object.
    /// </summary>
    /// <remarks>
    /// The source code has been taken from <a href="http://dbe.codeplex.com">Word Content Control Toolkit</a> and
    /// modified by the author of this class.
    /// </remarks>
    public partial class XMLStructureUserControl : UserControl
    {
        /// <summary>
        /// The key for the <i>Expanded</i> image in the <see cref="ImageList"/>.
        /// </summary>
        private const string ImageKeyXmlElementExpanded = "xmlElementExpanded";

        /// <summary>
        /// The key for the <i>Collapsed</i> image in the <see cref="ImageList"/>.
        /// </summary>
        private const string ImageKeyXmlElementCollapsed = "xmlElementCollapsed";

        /// <summary>
        /// The key for the <i>Element Data</i> image in the <see cref="ImageList"/>.
        /// </summary>
        private const string ImageKeyXmlElementData = "xmlElementData";

        /// <summary>
        /// The key for the <i>Attribute</i> image in the <see cref="ImageList"/>.
        /// </summary>
        private const string ImageKeyXmlAttribute = "xmlAttribute";

        /// <summary>
        /// The key for the <i>Comment</i> image in the <see cref="ImageList"/>.
        /// </summary>
        private const string ImageKeyXmlComment = "xmlComment";

        /// <summary>
        /// Used to build a string representation for a XPath prefix mapping.
        /// </summary>
        private readonly XPathPrefixMapping xpathPrefixMapping;

        /// <summary>
        /// Initializes a new instance of the <see cref="XMLStructureUserControl"/> class.
        /// </summary>
        public XMLStructureUserControl()
        {
            this.InitializeComponent();
            this.xpathPrefixMapping = new XPathPrefixMapping();
        }

        /// <summary>
        /// Handles the event that occurs after the selected <see cref="TreeNode"/> has changed.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event data.</param>
        public delegate void SelectedTreeNodeChangedHandler(
            object sender,
            XPathEventArgs e);

        /// <summary>
        /// Raised after the selected <see cref="TreeNode"/> has changed.
        /// </summary>
        public event EventHandler<XPathEventArgs> SelectedTreeNodeChanged;

        /// <summary>
        /// Gets the XPath event arguments for the currently selected <see cref="XmlNode"/>.
        /// </summary>
        public XPathEventArgs XPathEventArgs { get; private set; }

        /// <summary>
        /// Gets or sets the root <see cref="XmlNode"/>.
        /// </summary>
        private XmlNode XmlNodeRoot { get; set; }

        /// <summary>
        /// Gets or sets the root <see cref="TreeNode"/>.
        /// </summary>
        private TreeNode TreeNodeRoot { get; set; }

        /// <summary>
        /// Handles the <c>SelectedXmlDocumentChanged</c> event of this <see cref="XMLStructureUserControl"/>.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event data.</param>
        internal void OnXmlDocumentSelected(object sender, SelectedXmlDocumentChangedEventArgs e)
        {
            this.xpathPrefixMapping.Clear();
            this.PopulateTreeView(e.XmlDocument.DocumentElement);
            this.treeViewStructure.Enabled = true;
            this.TreeNodeRoot.ExpandTreeNode(2);
            this.ScrollTreeViewToTheTop();
            this.toolStrip.Enabled = true;
        }

        /// <summary>
        /// Populates the <see cref="TreeView"/> with the data from the specified <see cref="XmlNode"/>.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> whose data is used to populate the <see cref="TreeView"/>.</param>
        private void PopulateTreeView(XmlNode xmlNode)
        {
            this.XmlNodeRoot = xmlNode;
            this.TreeNodeRoot = this.CreateTreeNodeForXmlNode(this.XmlNodeRoot);
            this.TreeNodeRoot.Tag = this.XmlNodeRoot;

            this.treeViewStructure.Nodes.Clear();
            this.treeViewStructure.Nodes.Add(this.TreeNodeRoot);
            this.PopulateTreeNode(this.XmlNodeRoot, this.TreeNodeRoot);      
        }

        /// <summary>
        /// Populates the specified <see cref="TreeNode"/> with the data from the specified <see cref="XmlNode"/>.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> which is used as the tag for the <see cref="TreeNode"/>.</param>
        /// <param name="treeNode">The <see cref="TreeNode"/> to populate.</param>
        private void PopulateTreeNode(XmlNode xmlNode, TreeNode treeNode)
        {
            if (null == xmlNode)
            {
                throw new ArgumentNullException("xmlNode");
            }

            if (null == treeNode)
            {
                return;
            }

            TreeNode treeNodeChild = null;

            foreach (XmlNode xmlChildNode in xmlNode.ChildNodes)
            {
                treeNodeChild = this.CreateTreeNodeForXmlNode(xmlChildNode);
                
                if (null != treeNodeChild)
                {
                    treeNode.Nodes.Add(treeNodeChild);
                }

                this.PopulateTreeNode(xmlChildNode, treeNodeChild);

                // TODO This doesnt't belong into this method (separation of concerns).
                this.xpathPrefixMapping.AddNamespace(xmlChildNode.NamespaceURI);
            }

            if (null != xmlNode.Attributes)
            {
                foreach (XmlNode xmlNodeChild in xmlNode.Attributes)
                {
                    treeNodeChild = this.CreateTreeNodeForXmlNode(xmlNodeChild);

                    if (null != treeNodeChild)
                    {
                        treeNode.Nodes.Add(treeNodeChild);
                    }

                    this.PopulateTreeNode(xmlNodeChild, treeNodeChild);
                }
            }
        }

        /// <summary>
        /// Creates a new <see cref="TreeNode"/> for the specified <see cref="XmlNode"/>.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> which is used as the tag for the <see cref="TreeNode"/>.</param>
        /// <returns>The newly created <see cref="TreeNode"/>.</returns>
        private TreeNode CreateTreeNodeForXmlNode(XmlNode xmlNode)
        {
            TreeNode result = new TreeNode();

            // TODO Apply "Replace Conditional with Polymorphism".
            switch (xmlNode.NodeType)
            {
                case XmlNodeType.Element:
                    result.Text = xmlNode.LocalName;
                    string imageKey = xmlNode.IsLeaf()
                        ? ImageKeyXmlElementData
                        : ImageKeyXmlElementCollapsed;
                    this.SetImageKeysForTreeNode(result, imageKey);
                    break;
                case XmlNodeType.Attribute:
                    result.Text = xmlNode.LocalName;
                    this.SetImageKeysForTreeNode(result, ImageKeyXmlAttribute);
                    break;
                case XmlNodeType.Comment:
                    result.Text = "<!--" + xmlNode.Value + "-->";
                    this.SetImageKeysForTreeNode(result, ImageKeyXmlComment);
                    result.ForeColor = Color.DarkGreen;
                    break;
                case XmlNodeType.Document:
                    result.Text = "/";
                    this.SetImageKeysForTreeNode(result, ImageKeyXmlElementCollapsed);
                    break;
                default:
                    result = null;
                    break;
            }

            if (null != result)
            {
                result.Tag = xmlNode;
            }

            return result;
        }

        /// <summary>
        /// Sets the image keys of the <see cref="TreeNode"/> to the specified image key.
        /// </summary>
        /// <param name="treeNode">The <see cref="TreeNode"/> to modify.</param>
        /// <param name="imageKey">The image key.</param>
        private void SetImageKeysForTreeNode(TreeNode treeNode, string imageKey)
        {
            treeNode.ImageKey = imageKey;
            treeNode.SelectedImageKey = imageKey;
        }

        /// <summary>
        /// Scrolls the <see cref="TreeView"/> to the top.
        /// </summary>
        private void ScrollTreeViewToTheTop()
        {
            this.treeViewStructure.Nodes[0].EnsureVisible();
        }

        /// <summary>
        /// Handles the event that occurs after a <see cref="TreeNode"/> is expanded.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnAfterExpand(object sender, TreeViewEventArgs e)
        {
            XmlNode xmlNode = (XmlNode)e.Node.Tag;

            if (!xmlNode.IsLeaf())
            {
                this.SetImageKeysForTreeNode(e.Node, ImageKeyXmlElementExpanded);
            }
        }

        /// <summary>
        /// Handles the event that occurs after a <see cref="TreeNode"/> is collapsed.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnAfterCollapse(object sender, TreeViewEventArgs e)
        {
            XmlNode xmlNode = (XmlNode)e.Node.Tag;

            if (!xmlNode.IsLeaf())
            {
                this.SetImageKeysForTreeNode(e.Node, ImageKeyXmlElementCollapsed);
            }
        }

        /// <summary>
        /// Handles the event that occurs after a <see cref="TreeNode"/> is selected.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnAfterSelect(object sender, TreeViewEventArgs e)
        {
            XmlNode xmlNode = (XmlNode)e.Node.Tag;

            // Create the event arguments.
            this.XPathEventArgs = new XPathEventArgs(xmlNode.XPathExpression(this.xpathPrefixMapping.DefaultNamespaceURIs), this.xpathPrefixMapping.ToString());

            // Raise the event.
            this.SelectedTreeNodeChanged(this, this.XPathEventArgs);

            const string NamespacePrefix = "xmlns";

            bool isNamespaceNode = xmlNode.NodeType == XmlNodeType.Attribute && (NamespacePrefix == xmlNode.LocalName
                || NamespacePrefix == xmlNode.Prefix);
            this.toolStripButtonSelectNode.Enabled = !isNamespaceNode;
        }

        /// <summary>
        /// Handles the event that occurs when the input focus leaves the <see cref="TreeView"/> control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnLeaveTreeViewStructure(object sender, EventArgs e)
        {
            this.toolStripButtonSelectNode.Enabled = false;
        }

        /// <summary>
        /// Handles the event that occurs when the <see cref="TreeView"/> control is entered.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnEnterTreeViewStructure(object sender, EventArgs e)
        {
            this.toolStripButtonSelectNode.Enabled = null != this.treeViewStructure.SelectedNode;
        }

        /// <summary>
        /// Handles the event that occurs when the <see cref="ToolStripButton"/> <i>Expand Node</i> is clicked.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnClickToolStripButtonExpandNode(object sender, EventArgs e)
        {
            this.TreeNodeRoot.ExpandTreeNode();
            this.ScrollTreeViewToTheTop();
        }

        /// <summary>
        /// Handles the event that occurs when the <see cref="ToolStripButton"/> <i>Collapse Node</i> is clicked.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnClickToolStripButtonCollapseNode(object sender, EventArgs e)
        {
            this.TreeNodeRoot.CollapseTreeNode();
            this.ScrollTreeViewToTheTop();
        }

        /// <summary>
        /// Handles the event that occurs when the <see cref="ToolStripButton"/> <i>Select Node</i> is clicked.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void OnClickToolStripButtonSelectNode(object sender, EventArgs e)
        {
            // TODO I am not sure if this is a good approach. An alternative is
            // to change the visibility of the ToolStripButton to public and
            // allow the Form to register a click event. 
            Form parent = (Form)this.Parent;
            parent.DialogResult = DialogResult.OK;
            parent.Close();
        }
    }
}
