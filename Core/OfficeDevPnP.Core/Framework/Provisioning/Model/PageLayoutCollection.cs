namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of PageLayout objects
    /// </summary>
    public partial class PageLayoutCollection : ProvisioningTemplateCollection<PageLayout>
    {
        /// <summary>
        /// Constructor for PageLayoutCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public PageLayoutCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
