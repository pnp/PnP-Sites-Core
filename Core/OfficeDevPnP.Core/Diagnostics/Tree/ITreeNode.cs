using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    /// <summary>
    /// Contains Tree node properties and methods
    /// </summary>
    /// <typeparam name="T">Generic type</typeparam>
    public interface ITreeNode<T> : ITreeNode
    {
        /// <summary>
        /// Top root node
        /// </summary>
        ITreeNode<T> Root { get; }

        /// <summary>
        /// Parent node
        /// </summary>
        ITreeNode<T> Parent { get; set; }

        /// <summary>
        /// Sets parent node to the child nodes
        /// </summary>
        /// <param name="Node">Node to which we need to set parent node</param>
        /// <param name="UpdateChildNodes">Updates child nodes</param>
        void SetParent(ITreeNode<T> Node, bool UpdateChildNodes = true);

        /// <summary>
        /// Generic type value
        /// </summary>
        T Value { get; set; }

        /// <summary>
        /// Gets child node
        /// </summary>
        TreeNodeList<T> Children { get; }
    }

    /// <summary>
    /// Contains tree node properties
    /// </summary>
    public interface ITreeNode
    {
        /// <summary>
        /// All nodes along path toward root: Parent, Parent.Parent, Parent.Parent.Parent, ...
        /// </summary>
        IEnumerable<ITreeNode> Ancestors { get; }

        /// <summary>
        /// Parent node
        /// </summary>
        ITreeNode ParentNode { get; }

        /// <summary>
        /// Direct descendants
        /// </summary>
        IEnumerable<ITreeNode> ChildNodes { get; }

        /// <summary>
        /// All Child nodes. Children, Children[i].Children, ...
        /// </summary>
        IEnumerable<ITreeNode> Descendants { get; }

        /// <summary>
        /// Distance from Root
        /// </summary>
        int Depth { get; }
        //void OnDepthChanged();

        /// <summary>
        /// Distance from deepest descendant
        /// </summary>
        int Height { get; }
        //void OnHeightChanged();
    }
}
