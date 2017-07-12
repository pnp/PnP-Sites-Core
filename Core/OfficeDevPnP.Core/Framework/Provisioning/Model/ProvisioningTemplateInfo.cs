using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Class holds the required Provisioning Template Information
    /// </summary>
    public partial class ProvisioningTemplateInfo
    {
        #region Private Properties

        private string _templateId;

        #endregion
        /// <summary>
        /// Gets or sets the template id for the provisioning template
        /// </summary>
        public string TemplateId { get { return _templateId; } set { _templateId = value; } }

        /// <summary>
        /// Gets or sets the template version for the provisioning template
        /// </summary>
        public Double TemplateVersion { get; set; }
        /// <summary>
        /// Gets or sets the template site policy for the provisioning template
        /// </summary>
        public string TemplateSitePolicy { get; set; }
        /// <summary>
        /// Gets or sets the provisioning time for the provisioning template
        /// </summary>
        public DateTime ProvisioningTime { get; set; }
        /// <summary>
        /// Gets or sets the result for the provisioning template
        /// </summary>
        public bool Result { get; set; }
    }
}
