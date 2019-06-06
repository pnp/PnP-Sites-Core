using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Provisioning Hierarchy children
    /// </summary>
    /// <typeparam name="T">The type of Item for the collection</typeparam>
    public abstract class BaseProvisioningHierarchyObjectCollection<T> : Collection<T>, IProvisioningHierarchyDescendant
        where T : BaseHierarchyModel
    {
        /// <summary>
        /// Custom constructor to manage the ParentTemplate for the collection 
        /// and all the children of the collection
        /// </summary>
        /// <param name="parentHierarchy">Parent provisioning template</param>
        public BaseProvisioningHierarchyObjectCollection(ProvisioningHierarchy parentHierarchy)
        {
            this._parentHierarchy = parentHierarchy;
        }

        private ProvisioningHierarchy _parentHierarchy;

        /// <summary>
        /// References the parent ProvisioningTemplate for the current provisioning artifact
        /// </summary>
        public virtual ProvisioningHierarchy ParentHierarchy
        {
            get
            {
                return (this._parentHierarchy);
            }
            internal set
            {
                this._parentHierarchy = value;
            }
        }

        protected override void InsertItem(int index, T item)
        {
            base.InsertItem(index, item);
            item.ParentHierarchy = this.ParentHierarchy;
        }

        protected override void RemoveItem(int index)
        {
            this.Items[index].ParentHierarchy = null;
            base.RemoveItem(index);
        }

        protected override void SetItem(int index, T item)
        {
            base.SetItem(index, item);
            item.ParentHierarchy = this.ParentHierarchy;
        }

        protected override void ClearItems()
        {
            foreach (var item in this.Items)
            {
                item.ParentHierarchy = null;
            }
            base.ClearItems();
        }

        /// <summary>
        /// Adds item to the collection
        /// </summary>
        /// <param name="collection">IEnumerable type object</param>
        public virtual void AddRange(IEnumerable<T> collection)
        {
            if (collection != null)
            {
                foreach (var item in collection)
                {
                    this.Add(item);
                }
            }
        }

        /// <summary>
        /// Finds an item matching a search predicate
        /// </summary>
        /// <param name="match">The matching predicate to use for finding any target item</param>
        /// <returns>The target item matching the find predicate</returns>
        /// <remarks>We implemented this to adhere to the generic List of T behavior</remarks>
        public T Find(Predicate<T> match)
        {
            return (this.FirstOrDefault(item => match(item)));
        }

        /// <summary>
        /// Finds index of the item matching the search predicate
        /// </summary>
        /// <param name="match">The matching predicate to use for finding any target item</param>
        /// <returns>Returns matching index</returns>
        public Int32 FindIndex(Predicate<T> match)
        {
            return (this.FindIndex(0, this.Count, match));
        }

        /// <summary>
        /// Finds index of the item matching the search predicate based on start index provided.
        /// </summary>
        /// <param name="startIndex">Starting index</param>
        /// <param name="match">The matching predicate to use for finding any target item</param>
        /// <returns>Returns matching index</returns>
        public int FindIndex(int startIndex, Predicate<T> match)
        {
            return (this.FindIndex(startIndex, this.Count - startIndex, match));
        }

        /// <summary>
        /// Finds index of the item matching the search predicate based on start index and count.
        /// </summary>
        /// <param name="startIndex">Starting index</param>
        /// <param name="count">Index Count</param>
        /// <param name="match">The matching predicate to use for finding any target item</param>
        /// <returns>Returns matching index</returns>
        public int FindIndex(int startIndex, int count, Predicate<T> match)
        {
            if (startIndex > this.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(startIndex));
            }
            if ((count < 0) || (startIndex > (this.Count - count)))
            {
                throw new ArgumentOutOfRangeException(nameof(count));
            }
            if (match == null)
            {
                throw new ArgumentNullException(nameof(match));
            }

            int num = startIndex + count;
            for (int i = startIndex; i < num; i++)
            {
                if (match(this.Items[i]))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Removes all the matching indexes
        /// </summary>
        /// <param name="match">The matching predicate to use for removing target item</param>
        /// <returns>Returns count of the matches after removing</returns>
        public int RemoveAll(Predicate<T> match)
        {
            if (match == null)
            {
                throw new ArgumentNullException(nameof(match));
            }

            List<Int32> matches = new List<Int32>();

            for (Int32 index = 0; index < this.Items.Count; index++)
            {
                if (match(this.Items[index]))
                    matches.Add(index);
            }

            foreach (var index in matches.OrderByDescending(i => i))
            {
                this.Items.RemoveAt(index);
            }

            return (matches.Count());
        }
    }
}
