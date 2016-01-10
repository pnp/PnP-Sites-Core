namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of Field objects
    /// </summary>
    public partial class FieldCollection : ProvisioningTemplateCollection<Field>
    {
        public FieldCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {
        }
    }
}
