using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    public interface ITreeNode<T> : ITreeNode
    {
        ITreeNode<T> Root { get; }

        ITreeNode<T> Parent { get; set; }
        void SetParent(ITreeNode<T> Node, bool UpdateChildNodes = true);

        T Value { get; set; }

        TreeNodeList<T> Children { get; }
    }

    public interface ITreeNode
    {
        // all nodes along path toward root: Parent, Parent.Parent, Parent.Parent.Parent, ...
        IEnumerable<ITreeNode> Ancestors { get; }

        ITreeNode ParentNode { get; }

        // direct descendants
        IEnumerable<ITreeNode> ChildNodes { get; }

        // Children, Children[i].Children, ...
        IEnumerable<ITreeNode> Descendants { get; }

        // distance from Root
        int Depth { get; }
        //void OnDepthChanged();

        // distance from deepest descendant
        int Height { get; }
        //void OnHeightChanged();
    }
}
