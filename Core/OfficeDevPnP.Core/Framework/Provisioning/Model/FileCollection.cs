namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of File objects
    /// </summary>
    public partial class FileCollection : ProvisioningTemplateCollection<File>
    {
        /// <summary>
        /// Constructor for FileCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public FileCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
