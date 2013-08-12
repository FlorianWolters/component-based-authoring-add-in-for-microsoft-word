//------------------------------------------------------------------------------
// <copyright file="TreeNodeExtensions.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.Extensions
{
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="TreeNodeExtensions"/> contains extension methods for the class <see cref="TreeNode"/>.
    /// </summary>
    public static class TreeNodeExtensions
    {
        /// <summary>
        /// Expands the specified <see cref="TreeNode"/> once.
        /// </summary>
        /// <param name="treeNode">The <see cref="TreeNode"/> to expand.</param>
        public static void ExpandTreeNode(this TreeNode treeNode)
        {
            if (treeNode.IsExpanded)
            {
                foreach (TreeNode treeNodeChild in treeNode.Nodes)
                {
                    treeNodeChild.ExpandTreeNode();
                }
            }
            else
            {
                // Expand the TreeNode if it is not expanded.
                treeNode.Expand();
            }
        }

        /// <summary>
        /// Collapses the specified <see cref="TreeNode"/> once.
        /// </summary>
        /// <param name="treeNode">The <see cref="TreeNode"/> to collapse.</param>
        public static void CollapseTreeNode(this TreeNode treeNode)
        {
            if (CountExpandedTreeNodesBelow(treeNode) > 0)
            {
                foreach (TreeNode treeNodeChild in treeNode.Nodes)
                {
                    treeNodeChild.CollapseTreeNode();
                }
            }
            else
            {
                // Collapse the TreeNode if there are no expanded TreeNodes below.
                treeNode.Collapse();
            }
        }

        /// <summary>
        /// Expands the specified <see cref="TreeNode"/> for the specified number of times.
        /// </summary>
        /// <param name="treeNode">The <see cref="TreeNode"/> to expand.</param>
        /// <param name="times">The number of time to expand the <see cref="TreeNode"/>.</param>
        public static void ExpandTreeNode(this TreeNode treeNode, int times)
        {
            for (int i = 0; i < times; ++i)
            {
                treeNode.ExpandTreeNode();
            }
        }

        /// <summary>
        /// Returns the number of expanded <see cref="TreeNode"/>s below the
        /// specified <see cref="TreeNode"/>.
        /// </summary>
        /// <param name="treeNode">The <see cref="TreeNode"/> whose expanded child <see cref="TreeNode"/>s to count.</param>
        /// <returns>The number of expanded child <see cref="TreeNode"/>s for the specified <see cref="TreeNode"/>.</returns>
        public static int CountExpandedTreeNodesBelow(this TreeNode treeNode)
        {
            int result = 0;

            foreach (TreeNode treeNodeChild in treeNode.Nodes)
            {
                if (treeNodeChild.IsExpanded)
                {
                    ++result;
                }
            }

            return result;
        }

        /// <summary>
        /// Returns the number of collapses <see cref="TreeNode"/>s below the
        /// specified <see cref="TreeNode"/>.
        /// </summary>
        /// <param name="treeNode">The <see cref="TreeNode"/> whose collapsed child <see cref="TreeNode"/>s to count.</param>
        /// <returns>The number of collapsed child <see cref="TreeNode"/>s for the specified <see cref="TreeNode"/>.</returns>
        public static int CountCollapsedTreeNodesBelow(this TreeNode treeNode)
        {
            int result = 0;

            foreach (TreeNode treeNodeChild in treeNode.Nodes)
            {
                if (!treeNodeChild.IsExpanded)
                {
                    ++result;
                }
            }

            return result;
        }
    }
}
