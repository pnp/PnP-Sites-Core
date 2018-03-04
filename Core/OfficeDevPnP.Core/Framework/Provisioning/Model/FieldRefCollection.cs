namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of FieldRef objects
    /// </summary>
    public partial class FieldRefCollection : ProvisioningTemplateCollection<FieldRef>
    {
        /// <summary>
        /// Constructor for FieldRefCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public FieldRefCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
