namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of StorageEntity objects
    /// </summary>
    public partial class StorageEntityCollection : ProvisioningTemplateCollection<StorageEntity>
    {
        /// <summary>
        /// Constructor for StorageEntityCollection
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public StorageEntityCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
