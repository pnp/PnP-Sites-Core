using System.Collections.Generic;
using System.ComponentModel;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{

    public interface ITreeNodeList<T> : IList<ITreeNode<T>>
    {
        new ITreeNode<T> Add(ITreeNode<T> node);
    }
}
