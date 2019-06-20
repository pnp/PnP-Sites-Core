using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    /// <summary>
    /// Defines type of tree travarsal
    /// </summary>
    public enum UpDownTraversalType
    {
        /// <summary>
        /// Top to down tree traversal
        /// </summary>
        TopDown,
        /// <summary>
        /// Bottom to up tree traversal
        /// </summary>
        BottomUp
    }

    /// <summary>
    /// Defines type of tree traversal (Depth/Breadth)
    /// </summary>
    public enum DepthBreadthTraversalType
    {
        /// <summary>
        /// Depth first tree traversal
        /// </summary>
        DepthFirst,
        /// <summary>
        /// Breadth first traversal
        /// </summary>
        BreadthFirst
    }

    /// <summary>
    /// Defines changes done on node
    /// </summary>
    public enum NodeChangeType
    {
        /// <summary>
        /// Node is added
        /// </summary>
        NodeAdded,
        /// <summary>
        /// Node is removed
        /// </summary>
        NodeRemoved
    }

    /// <summary>
    /// Defines type of node ralation
    /// </summary>
    public enum NodeRelationType
    {
        /// <summary>
        /// Ancestor node
        /// </summary>
        Ancestor,
        /// <summary>
        /// Parent Node
        /// </summary>
        Parent,
        /// <summary>
        /// Self Node
        /// </summary>
        Self,
        /// <summary>
        /// Child Node
        /// </summary>
        Child,
        /// <summary>
        /// Descendant Node
        /// </summary>
        Descendant
    }
}
