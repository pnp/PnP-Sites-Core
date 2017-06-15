using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Collection of SupportedUILanguage objects
    /// </summary>
    public partial class SupportedUILanguageCollection : ProvisioningTemplateCollection<SupportedUILanguage>
    {
        /// <summary>
        /// Constructor for SupportedUILangaugeCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public SupportedUILanguageCollection(ProvisioningTemplate parentTemplate) : base(parentTemplate)
        {

        }
    }
}
