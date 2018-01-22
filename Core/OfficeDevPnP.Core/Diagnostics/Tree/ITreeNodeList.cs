using System.Collections.Generic;
using System.ComponentModel;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    /// <summary>
    /// Contains add mothod to add node to a tree
    /// </summary>
    /// <typeparam name="T">Generic type tree node</typeparam>
    public interface ITreeNodeList<T> : IList<ITreeNode<T>>
    {
        /// <summary>
        /// Adds node to a tree
        /// </summary>
        /// <param name="node">Tree node to add to a tree</param>
        /// <returns>Returns Generic type ITreeNode object</returns>
        new ITreeNode<T> Add(ITreeNode<T> node);
    }
}
