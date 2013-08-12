//------------------------------------------------------------------------------
// <copyright file="XPathUserControl.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.UserControls
{
    using System.Windows.Forms;
    using FlorianWolters.Windows.Forms.XML.EventArgs;

    /// <summary>
    /// The class <see cref="XPathUserControl"/> implements a <see cref="UserControl"/> which allows to display the
    /// XPath expression and XPath namespace-mappings represented via a <see cref="XPathEventArgs"/> object.
    /// </summary>
    /// <remarks>
    /// The source code has been taken from <a href="http://dbe.codeplex.com">Word Content Control Toolkit</a> and
    /// modified by the author of this class.
    /// </remarks>
    public partial class XPathUserControl : UserControl
    {
        /// <summary>
        /// The (invalid) XPath expression for a XML attribute which specifies a namespace URI.
        /// </summary>
        private const string NamespaceAttributeXPathExpression = "/@xmlns";

        /// <summary>
        /// Initializes a new instance of the <see cref="XPathUserControl"/> class.
        /// </summary>
        public XPathUserControl()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Gets the XPath expression.
        /// </summary>
        public string XPathExpression
        {
            get
            {
                return this.textBoxXPathExpression.Text;
            }
        }

        /// <summary>
        /// Handles the event that occurs after a <see cref="TreeNode"/> is selected in the <see cref="TreeView"/>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        internal void OnAfterSelectTreeView(object sender, XPathEventArgs e)
        {
            if (!e.XPathExpression.Contains(NamespaceAttributeXPathExpression))
            {
                this.textBoxXPathExpression.Text = e.XPathExpression;
                this.textBoxXPathPrefixMapping.Text = e.XPathPrefixMapping;
            }
            else
            {
                this.textBoxXPathExpression.Text = string.Empty;
                this.textBoxXPathPrefixMapping.Text = string.Empty;
            }
        }
    }
}
