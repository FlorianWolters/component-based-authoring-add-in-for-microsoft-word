//------------------------------------------------------------------------------
// <copyright file="XmlNodeExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Xml;

    /// <summary>
    /// The class <see cref="XmlNodeExtensions"/> contains extension methods for the class <see cref="XmlNode"/>.
    /// </summary>
    /// <remarks>
    /// The source code has been taken from <a href="http://dbe.codeplex.com">Word Content Control Toolkit</a> and
    /// modified by the author of this class.
    /// </remarks>
    public static class XmlNodeExtensions
    {
        /// <summary>
        /// Determines whether the specified <see cref="XmlNode"/> is a leaf.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> to check.</param>
        /// <returns><c>true</c> if the specified <see cref="XmlNode"/> is a leaf; <c>false</c> otherwise.</returns>
        public static bool IsLeaf(this XmlNode xmlNode)
        {
            bool result = true;

            foreach (XmlNode xmlChildNode in xmlNode.ChildNodes)
            {
                if (xmlChildNode.NodeType == XmlNodeType.Element)
                {
                    result = false;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Returns a XPath expression for the specified <see cref="XmlNode"/> and the specified namespace URIs.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> to analyze.</param>
        /// <param name="availableNamespaces">The available namespace URIs in the XML node.</param>
        /// <returns>A XPath expression for the specified <see cref="XmlNode"/>.</returns>
        public static string XPathExpression(this XmlNode xmlNode, IList<string> availableNamespaces)
        {
            string strThis = null;
            string strThisName = string.Empty;
            string namespaceURI = string.Empty;
            XmlNode xmlParentNode = xmlNode.ParentNode;
            bool checkParent = true;
            
            // TODO use StringBuilder.
            switch (xmlNode.NodeType)
            {
                case XmlNodeType.Element:
                    strThisName = xmlNode.LocalName;

                    if (string.Empty != xmlNode.NamespaceURI)
                    {
                        int namespacePos = availableNamespaces.IndexOf(xmlNode.NamespaceURI);

                        if (-1 == namespacePos)
                        {
                            throw new ArgumentException("availableNamespaces");
                        }

                        namespaceURI = "ns" + namespacePos + ":";
                    }

                    strThis = "/" + namespaceURI + strThisName + XPathPositionStringFromXmlNode(xmlNode);
                    break;
                case XmlNodeType.Attribute:
                    strThisName = xmlNode.Name;
                    strThis = "/@" + strThisName;

                    if (strThisName != "xmlns")
                    {
                        xmlParentNode = xmlNode.SelectSingleNode("..", null);
                    }
                    else
                    {
                        checkParent = false;
                    }

                    break;
                case XmlNodeType.ProcessingInstruction:
                    XmlProcessingInstruction xpi = xmlNode as XmlProcessingInstruction;
                    strThis = "/processing-instruction(";
                    strThis = strThis + ")" + XPathPositionStringFromXmlNode(xmlNode);
                    break;
                case XmlNodeType.Text:
                    strThis = "/text()" + XPathPositionStringFromXmlNode(xmlNode);
                    break;
                case XmlNodeType.Comment:
                    strThis = "/comment()" + XPathPositionStringFromXmlNode(xmlNode);
                    break;
                case XmlNodeType.Document:
                    strThis = string.Empty;
                    checkParent = false;
                    break;
                case XmlNodeType.EntityReference:
                case XmlNodeType.CDATA:
                    strThis = "/text()" + XPathPositionStringFromXmlNode(xmlNode);
                    break;
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    break;
            }

            return strThis.Insert(0, checkParent ? XPathExpression(xmlParentNode, availableNamespaces) : string.Empty);
        }

        /// <summary>
        /// Returns a XPath position substring for the specified <see cref="XmlNode"/>.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> to analyze.</param>
        /// <returns>The XPath position substring.</returns>
        private static string XPathPositionStringFromXmlNode(XmlNode xmlNode)
        {
            string result = string.Empty;

            if (ChildNodeCountForParentNode(xmlNode) > 0)
            {
                result = "[" + SiblingNodeCount(xmlNode).ToString() + "]";
            }

            return result;
        }

        /// <summary>
        /// Returns the number of child nodes for the parent node of the specified <see cref="XmlNode"/>.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> to analyze.</param>
        /// <returns>The number of child for the parent node of the specified <see cref="XmlNode"/>.</returns>
        private static long ChildNodeCountForParentNode(XmlNode xmlNode)
        {
            long childNodeCount = 0;

            if (xmlNode.ParentNode != null)
            {
                childNodeCount = xmlNode.ParentNode.ChildNodes.Count;
            }

            return childNodeCount;
        }

        /// <summary>
        /// Returns the number of similar sibling nodes for the specified <see cref="XmlNode"/>.
        /// </summary>
        /// <param name="xmlNode">The <see cref="XmlNode"/> to analyze.</param>
        /// <returns>The number of similar sibling nodes for the specified <see cref="XmlNode"/>.</returns>
        private static long SiblingNodeCount(XmlNode xmlNode)
        {
            XmlNode xmlNodePreviousSibling = null;
            long siblingNodeCount = 1;

            while (null != (xmlNodePreviousSibling = xmlNode.PreviousSibling))
            {
                if (xmlNodePreviousSibling.NodeType == xmlNode.NodeType
                    && xmlNodePreviousSibling.LocalName == xmlNode.LocalName
                    && xmlNodePreviousSibling.NamespaceURI == xmlNode.NamespaceURI)
                {
                    ++siblingNodeCount;
                }

                xmlNode = xmlNodePreviousSibling;
            }

            return siblingNodeCount;
        }
    }
}
