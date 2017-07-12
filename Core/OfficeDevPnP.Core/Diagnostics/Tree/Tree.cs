namespace OfficeDevPnP.Core.Diagnostics.Tree
{
    /// <summary>
    /// Defines Tree
    /// </summary>
    /// <typeparam name="T">Generic type</typeparam>
    public class Tree<T> : TreeNode<T>
        where T : new()
    {
        /// <summary>
        /// Default Constructor
        /// </summary>
        public Tree() { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="RootValue">Generic type paramerter for root value of tree</param>
        public Tree(T RootValue)
        {
            Value = RootValue;
        }
    }
}
