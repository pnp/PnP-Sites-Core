namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SiteDesign objects
    /// </summary>
    public partial class SiteDesignCollection : BaseProvisioningTemplateObjectCollection<SiteDesign>
    {
        /// <summary>
        /// Constructor for SiteDesignCollection
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public SiteDesignCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
