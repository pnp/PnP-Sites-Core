namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Feature objects
    /// </summary>
    public partial class FeatureCollection : ProvisioningTemplateCollection<Feature>
    {
        public FeatureCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
