namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Page objects
    /// </summary>
    public partial class PageCollection : ProvisioningTemplateCollection<Page>
    {
        public PageCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
