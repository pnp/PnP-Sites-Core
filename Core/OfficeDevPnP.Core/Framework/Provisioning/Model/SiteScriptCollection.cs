namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SiteScript objects
    /// </summary>
    public partial class SiteScriptCollection : ProvisioningTemplateCollection<SiteScript>
    {
        /// <summary>
        /// Constructor for SiteScriptCollection
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public SiteScriptCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
