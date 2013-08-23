//------------------------------------------------------------------------------
// <copyright file="CustomXMLNodeExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Extensions
{
    using System.Collections.Generic;
    using System.Linq;
    using Office = Microsoft.Office.Core;

    /// <summary>
    /// The static class <see cref="CustomXMLNodeExtensions"/> contains extension methods for a custom XML node of a
    /// custom XML part, represented by an object of the class <see cref="Office.CustomXMLNode"/>.
    /// </summary>
    public static class CustomXMLNodeExtensions
    {
        /// <summary>
        /// Returns the level of the specified <see cref="Office.CustomXMLNode"/>.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to check.</param>
        /// <returns>The level.</returns>
        public static int GetLevel(this Office.CustomXMLNode node)
        {
            int result = 0;

            while (null != (node = node.ParentNode))
            {
                ++result;
            }

            return result;
        }

        /// <summary>
        /// Checks whether the specified <see cref="Office.CustomXMLNode"/> is an attribute.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to check.</param>
        /// <returns>
        /// <c>true</c> if the specified <see cref="Office.CustomXMLNode"/> is an attribute; <c>false</c> otherwise.
        /// </returns>
        public static bool IsAttribute(this Office.CustomXMLNode node)
        {
            return Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute == node.NodeType;
        }

        /// <summary>
        /// Checks whether the specified <see cref="Office.CustomXMLNode"/> is an element.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to check.</param>
        /// <returns>
        /// <c>true</c> if the specified <see cref="Office.CustomXMLNode"/> is an element; <c>false</c> otherwise.
        /// </returns>
        public static bool IsElement(this Office.CustomXMLNode node)
        {
            return Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement == node.NodeType;
        }

        /// <summary>
        /// Checks whether the specified <see cref="Office.CustomXMLNode"/> is a leaf element.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to check.</param>
        /// <returns>
        /// <c>true</c> if the specified <see cref="Office.CustomXMLNode"/> is a leaf element; <c>false</c> otherwise.
        /// </returns>
        public static bool IsLeafElement(this Office.CustomXMLNode node)
        {
            return IsElement(node) && 1 == node.ChildNodes.Count;
        }

        /// <summary>
        /// Checks whether the specified <see cref="Office.CustomXMLNode"/> is an empty element.
        /// </summary>
        /// <param name="node">The <see cref="Office.CustomXMLNode"/> to check.</param>
        /// <returns>
        /// <c>true</c> if the specified <see cref="Office.CustomXMLNode"/> is an empty element; <c>false</c> otherwise.
        /// </returns>
        public static bool IsEmptyElement(this Office.CustomXMLNode node)
        {
            return IsElement(node) && !node.HasChildNodes();
        }
    }
}
