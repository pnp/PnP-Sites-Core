namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of File objects
    /// </summary>
    public partial class FileCollection : ProvisioningTemplateCollection<File>
    {
        public FileCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
