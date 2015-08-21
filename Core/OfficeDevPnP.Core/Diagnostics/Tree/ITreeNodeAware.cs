namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    public interface ITreeNodeAware<T>
         where T : new()
    {
        TreeNode<T> Node { get; set; }
    }
}
