namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of PropertyBagEntry objects
    /// </summary>
    public partial class PropertyBagEntryCollection : ProvisioningTemplateCollection<PropertyBagEntry>
    {
        /// <summary>
        /// Constructor for PropertyBagEntryCollection
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public PropertyBagEntryCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
