namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class SharedFieldCollection : BaseProvisioningTemplateObjectCollection<SharedField>
    {
        /// <summary>
        /// Constructor for SharedFieldCollection class.
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public SharedFieldCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
