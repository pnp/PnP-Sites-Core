namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of PropertyBagEntry objects
    /// </summary>
    public partial class PropertyBagEntryCollection : ProvisioningTemplateCollection<PropertyBagEntry>
    {
        public PropertyBagEntryCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
