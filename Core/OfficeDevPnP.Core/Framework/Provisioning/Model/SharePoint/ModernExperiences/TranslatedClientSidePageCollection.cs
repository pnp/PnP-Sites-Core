using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// A collection of TranslatedClientSidePage objects
    /// </summary>
    public partial class TranslatedClientSidePageCollection : BaseProvisioningTemplateObjectCollection<TranslatedClientSidePage>
    {
        /// <summary>
        /// Constructor for TranslatedClientSidePageCollection class
        /// </summary>
        /// <param name="parentTemplate">Parent provisioning template</param>
        public TranslatedClientSidePageCollection(ProvisioningTemplate parentTemplate) :
            base(parentTemplate)
        {
        }
    }
}
