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
            foreach (var item in collection)
            {
                this.Add(item);
            }
        }
    }
}
