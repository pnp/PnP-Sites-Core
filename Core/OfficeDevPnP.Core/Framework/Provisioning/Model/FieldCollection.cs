namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Field objects
    /// </summary>
    public partial class FieldCollection : BaseProvisioningTemplateObjectCollection<Field>
    {
        /// <summary>
        /// Constructor for FieldCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public FieldCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {
        }
    }
}
