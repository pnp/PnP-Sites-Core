using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    public class TreeNode<T> : ITreeNode<T>
        where T : new()
    {
        public TreeNode()
        {
            _Parent = null;
            _ChildNodes = new TreeNodeList<T>(this);
        }

        public TreeNode(T Value)
        {
            this.Value = Value;
            _Parent = null;
            _ChildNodes = new TreeNodeList<T>(this);
        }

        public TreeNode(T Value, TreeNode<T> Parent)
        {
            this.Value = Value;
            _Parent = Parent;
            _ChildNodes = new TreeNodeList<T>(this);
        }

        public ITreeNode ParentNode
        {
            get { return _Parent; }
        }

        private ITreeNode<T> _Parent;
        public ITreeNode<T> Parent
        {
            get { return _Parent; }
            set { SetParent(value, true); }
        }

        public void SetParent(ITreeNode<T> node, bool updateChildNodes = true)
        {
            if (node == Parent)
                return;

            var oldParent = Parent;
            var oldParentHeight = Parent != null ? Parent.Height : 0;
            var oldDepth = Depth;

            // if oldParent isn't null
            // remove this node from its newly ex-parent's children
            if (oldParent != null && oldParent.Children.Contains(this))
                oldParent.Children.Remove(this, updateParent: false);

            // update the backing field
            _Parent = node;

            // add this node to its new parent's children
            if (_Parent != null && updateChildNodes)
                _Parent.Children.Add(this, updateParent: false);

            // if this operation has changed the height of any parent, initiate the bubble-up height changed event
            if (Parent != null)
            {
                var newParentHeight = Parent != null ? Parent.Height : 0;
            }
        }

        public ITreeNode<T> Root
        {
            get { return (Parent == null) ? this : Parent.Root; }
        }

        private TreeNodeList<T> _ChildNodes;
        public TreeNodeList<T> Children
        {
            get { return _ChildNodes; }
        }

        // non-generic iterator for interface-based support (From TreeNodeView, for example)
        public IEnumerable<ITreeNode> ChildNodes
        {
            get
            {
                foreach (ITreeNode node in Children)
                    yield return node;

                yield break;
            }
        }

        public IEnumerable<ITreeNode> Descendants
        {
            get
            {
                foreach (ITreeNode node in ChildNodes)
                {
                    yield return node;

                    foreach (ITreeNode descendant in node.Descendants)
                        yield return descendant;
                }

                yield break;
            }
        }

        public IEnumerable<ITreeNode> Subtree
        {
            get
            {
                yield return this;

                foreach (ITreeNode node in Descendants)
                    yield return node;

                yield break;
            }
        }

        public IEnumerable<ITreeNode> Ancestors
        {
            get
            {
                if (Parent == null)
                    yield break;

                yield return Parent;

                foreach (ITreeNode node in Parent.Ancestors)
                    yield return node;

                yield break;
            }
        }

        // [recurse up] descending aggregate property
        public int Height
        {
            get { return Children.Count == 0 ? 0 : Children.Max(n => n.Height) + 1; }
        }

        private T _Value;
        public T Value
        {
            get { return _Value; }
            set
            {
                if (value == null && _Value == null)
                    return;

                if (value != null && _Value != null && value.Equals(_Value))
                    return;

                _Value = value;

                // set Node if it's ITreeNodeAware
                if (_Value != null && _Value is ITreeNodeAware<T>)
                    (_Value as ITreeNodeAware<T>).Node = this;
            }
        }

        // [recurse up] bubble up aggregate property
        public int Depth
        {
            get { return (Parent == null ? 0 : Parent.Depth + 1); }
        }

        //private UpDownTraversalType _DisposeTraversal = UpDownTraversalType.BottomUp;
        //public UpDownTraversalType DisposeTraversal
        //{
        //    get { return _DisposeTraversal; }
        //    set { _DisposeTraversal = value; }
        //}

        //private bool _IsDisposed;
        //public bool IsDisposed
        //{
        //    get { return _IsDisposed; }
        //}

        //protected virtual void Dispose(bool cleanUpNative)
        //{
        //    CheckDisposed();
        //    OnDisposing();

        //    // clean up contained objects (in Value property)
        //    if (Value is IDisposable)
        //    {
        //        if (DisposeTraversal == UpDownTraversalType.BottomUp)
        //            foreach (TreeNode<T> node in Children)
        //                node.Dispose();

        //        (Value as IDisposable).Dispose();

        //        if (DisposeTraversal == UpDownTraversalType.TopDown)
        //            foreach (TreeNode<T> node in Children)
        //                node.Dispose();
        //    }

        //    _IsDisposed = true;

        //}

        //public virtual void Dispose()
        //{
        //    Dispose(false);
        //    GC.SuppressFinalize(this);
        //}

        //public event EventHandler Disposing;

        //protected void OnDisposing()
        //{
        //    if (Disposing != null)
        //        Disposing(this, EventArgs.Empty);
        //}

        //public void CheckDisposed()
        //{
        //    if (IsDisposed)
        //        throw new ObjectDisposedException(GetType().Name);
        //}

        public override string ToString()
        {
            return "Depth=" + Depth + ", Height=" + Height + ", Children=" + Children.Count;
        }
    }
}
