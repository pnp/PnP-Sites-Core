using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    /// <summary>
    /// Holds methods for Tree node
    /// </summary>
    /// <typeparam name="T">Generic type</typeparam>
    public class TreeNodeList<T> : List<ITreeNode<T>>, ITreeNodeList<T>, INotifyPropertyChanged
    {
        /// <summary>
        /// Gets or sets Parent node
        /// </summary>
        public ITreeNode<T> Parent { get; set; }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="parent">Generic type parent node</param>
        public TreeNodeList(ITreeNode<T> parent)
        {
            Parent = parent;
        }
        /// <summary>
        /// Adds node to a tree
        /// </summary>
        /// <param name="node">Node to be added to the tree</param>
        /// <returns>Returns TreeNode interface</returns>
        public new ITreeNode<T> Add(ITreeNode<T> node)
        {
            return Add(node, true);
        }

        protected internal ITreeNode<T> Add(ITreeNode<T> node, bool updateParent)
        {
            if (updateParent)
            {
                // force Node.SetParent to coordinate the update
                node.SetParent(Parent, UpdateChildNodes: true);
                return node;
            }

            base.Add(node);
            OnPropertyChanged("Count");
            return node;
        }
        /// <summary>
        /// Removes node from a tree
        /// </summary>
        /// <param name="node">Tree node to be removed from a tree</param>
        /// <returns>Returns status of Node removal from tree</returns>
        public new bool Remove(ITreeNode<T> node)
        {
            return Remove(node, true);
        }

        protected internal bool Remove(ITreeNode<T> node, bool updateParent)
        {
            if (node == null)
                throw new ArgumentNullException("node");

            // if we don't have it, we can't remove it
            if (!Contains(node))
                return false;

            if (updateParent)
            {
                // force Node.SetParent to coordinate the update
                node.SetParent(null, UpdateChildNodes: false);

                // we're successful if the node is no longer in the collection
                return !Contains(node);
            }

            var result = base.Remove(node);
            OnPropertyChanged("Count");
            return result;
        }
        /// <summary>
        /// Reprensets PropertyChangedEventHandler on a tree node
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string PropertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(PropertyName));
        }
        /// <summary>
        /// Retuns count of child nodes as a string
        /// </summary>
        /// <returns>Retuns count of child nodes as a string</returns>
        public override string ToString()
        {
            return "Count=" + Count;
        }
    }
}
