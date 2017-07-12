using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Generic keyed collection of items stored in the ProvisioningTemplate graph
    /// </summary>
    /// <typeparam name="TKey">The type of the Key for the keyed collection</typeparam>
    /// <typeparam name="TItem">The type of the Item for the keyed collection</typeparam>
    public abstract class ProvisioningTemplateDictionary<TKey, TItem> : KeyedCollection<TKey, TItem>, IProvisioningTemplateDescendant
        where TItem : BaseModel
    {
        /// <summary>
        /// Custom constructor to manage the ParentTemplate for the collection 
        /// and all the children of the collection
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public ProvisioningTemplateDictionary(ProvisioningTemplate parentTemplate)
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

        protected override void InsertItem(int index, TItem item)
        {
            base.InsertItem(index, item);
        }

        protected override void SetItem(int index, TItem item)
        {
            base.SetItem(index, item);
            item.ParentTemplate = this.ParentTemplate;
        }

        protected override void RemoveItem(int index)
        {
            this.Items[index].ParentTemplate = null;
            base.RemoveItem(index);
        }

        protected override void ClearItems()
        {
            foreach (var item in this.Items)
            {
                item.ParentTemplate = null;
            }
            base.ClearItems();
        }
    }
}
