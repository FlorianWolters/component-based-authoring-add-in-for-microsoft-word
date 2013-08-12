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

    public static class CustomXMLNodeExtensions
    {
        public static int GetLevel(this Office.CustomXMLNode node)
        {
            int result = 0;

            while (null != (node = node.ParentNode))
            {
                ++result;
            }

            return result;
        }

        public static bool IsAttribute(this Office.CustomXMLNode node)
        {
            return Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute == node.NodeType;
        }

        public static bool IsElement(this Office.CustomXMLNode node)
        {
            return Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement == node.NodeType;
        }

        public static bool IsLeafElement(this Office.CustomXMLNode node)
        {
            return IsElement(node) && 1 == node.ChildNodes.Count;
        }

        public static bool IsEmptyElement(this Office.CustomXMLNode node)
        {
            return IsElement(node) && !node.HasChildNodes();
        }

        public static IList<Office.CustomXMLNode> ToList(this Office.CustomXMLNodes nodes)
        {
            return new List<Office.CustomXMLNode>(nodes.Cast<Office.CustomXMLNode>());
        }
    }
}
