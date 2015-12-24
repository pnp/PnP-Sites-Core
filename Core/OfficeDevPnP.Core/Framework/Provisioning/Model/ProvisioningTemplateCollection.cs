using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Generic collection of items stored in the ProvisioningTemplate graph
    /// </summary>
    /// <typeparam name="T">The type of Item for the collection</typeparam>
    public abstract class ProvisioningTemplateCollection<T> : Collection<T>, IProvisioningTemplateDescendant
        where T : BaseModel
    {
        /// <summary>
        /// Custom constructor to manage the ParentTemplate for the collection 
        /// and all the children of the collection
        /// </summary>
        /// <param name="parentTemplate"></param>
        public ProvisioningTemplateCollection(ProvisioningTemplate parentTemplate)
        {
            this.ParentTemplate = parentTemplate;
        }

        private ProvisioningTemplate _parentTemplate;

        /// <summary>
        /// References the parent ProvisioningTemplate for the current provisioning artifact
        /// </summary>
        public virtual ProvisioningTemplate ParentTemplate
        {
            get
            {
                return (this._parentTemplate);
            }
            internal set
            {
                this._parentTemplate = value;
            }
        }

        protected override void InsertItem(int index, T item)
        {
            base.InsertItem(index, item);
            item.ParentTemplate = this.ParentTemplate;
        }

        protected override void RemoveItem(int index)
        {
            this.Items[index].ParentTemplate = null;
            base.RemoveItem(index);
        }

        protected override void SetItem(int index, T item)
        {
            base.SetItem(index, item);
            item.ParentTemplate = this.ParentTemplate;
        }

        protected override void ClearItems()
        {
            foreach (var item in this.Items)
            {
                item.ParentTemplate = null;
            }
            base.ClearItems();
        }

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
        public Int32 FindIndex(Predicate<T> match)
        {
            return (this.FindIndex(0, this.Count, match));
        }

        public int FindIndex(int startIndex, Predicate<T> match)
        {
            return (this.FindIndex(startIndex, this.Count - startIndex, match));
        }

        public int FindIndex(int startIndex, int count, Predicate<T> match)
        {
            if (startIndex > this.Count)
            {
                throw new ArgumentOutOfRangeException("startIndex");
            }
            if ((count < 0) || (startIndex > (this.Count - count)))
            {
                throw new ArgumentOutOfRangeException("count");
            }
            if (match == null)
            {
                throw new ArgumentNullException("match");
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

        public int RemoveAll(Predicate<T> match)
        {
            if (match == null)
            {
                throw new ArgumentNullException("match");
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
