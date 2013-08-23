//------------------------------------------------------------------------------
// <copyright file="XPathEventArgs.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.EventArgs
{
    using System;

    /// <summary>
    /// The class <see cref="XPathEventArgs"/> provides a XPath expression string and a XPath Namespace-Prefix mappings
    /// string as data for an event.
    /// </summary>
    public class XPathEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XPathEventArgs"/> class.
        /// </summary>
        /// <param name="xpathExpression">The XPath expression.</param>
        /// <param name="xpathPrefixMapping">The XPath Namespace-Prefix mappings.</param>
        public XPathEventArgs(string xpathExpression, string xpathPrefixMapping)
        {
            this.XPathExpression = xpathExpression;
            this.XPathPrefixMapping = xpathPrefixMapping;
        }

        /// <summary>
        /// Gets the XPath expression.
        /// </summary>
        public string XPathExpression { get; private set; }

        /// <summary>
        /// Gets the XPath Namespace-Prefix mappings.
        /// </summary>
        public string XPathPrefixMapping { get; private set; }
    }
}
