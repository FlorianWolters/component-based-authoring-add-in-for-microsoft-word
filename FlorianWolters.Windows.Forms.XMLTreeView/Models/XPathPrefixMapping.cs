//------------------------------------------------------------------------------
// <copyright file="XPathPrefixMapping.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.Models
{
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// The class <see cref="XPathPrefixMapping"/> allows so create a string which contains all Namespace-Prefix
    /// mappings (all mappings between prefixes and default namespace URIs) for a XPath expression.
    /// </summary>
    public class XPathPrefixMapping
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XPathPrefixMapping"/> class.
        /// </summary>
        public XPathPrefixMapping()
        {
            this.DefaultNamespaceURIs = new List<string>();
        }

        /// <summary>
        /// Gets the Namespace URIs for the default namespaces of this XPath prefix mapping.
        /// </summary>
        public IList<string> DefaultNamespaceURIs { get; private set; }

        /// <summary>
        /// Adds a Namespace URI as a default namespace for this XPath prefix mapping.
        /// </summary>
        /// <param name="namespaceURI">The namespace URI to add as a default namespace.</param>
        public void AddNamespace(string namespaceURI)
        {
            // Add unique namespaces only.
            if (!string.IsNullOrEmpty(namespaceURI) &&
                -1 == this.DefaultNamespaceURIs.IndexOf(namespaceURI))
            {
                this.DefaultNamespaceURIs.Insert(0, namespaceURI);
            }
        }

        /// <summary>
        /// Removed all namespace URIs of this XPath prefix mapping.
        /// </summary>
        public void Clear()
        {
            this.DefaultNamespaceURIs.Clear();
        }

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>A string that represents the current object.</returns>
        public override string ToString()
        {
            StringBuilder result = new StringBuilder();

            for (int i = 0; i < this.DefaultNamespaceURIs.Count; ++i)
            {
                result.Append("xmlns:ns");
                result.Append(i);
                result.Append("=\"");
                result.Append(this.DefaultNamespaceURIs[i]);
                result.Append("\" ");
            }

            return result.ToString().TrimEnd();
        }
    }
}
