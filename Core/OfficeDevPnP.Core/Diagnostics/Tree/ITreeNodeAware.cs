namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    /// <summary>
    /// Conints tree node properties
    /// </summary>
    /// <typeparam name="T">Generic type parameter</typeparam>
    public interface ITreeNodeAware<T>
         where T : new()
    {
        /// <summary>
        /// Sets or get TreeNode
        /// </summary>
        TreeNode<T> Node { get; set; }
    }
}
