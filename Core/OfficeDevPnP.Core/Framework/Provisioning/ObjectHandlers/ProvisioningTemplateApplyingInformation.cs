using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Privisioning Progress Delegate
    /// </summary>
    /// <param name="message">Provisioning message</param>
    /// <param name="step"></param>
    /// <param name="total">Handlers count</param>
    public delegate void ProvisioningProgressDelegate(string message, int step, int total);

    /// <summary>
    /// Provisioning Messages Delegate
    /// </summary>
    /// <param name="message">Provisioning message</param>
    /// <param name="messageType">Provisioning message type</param>
    public delegate void ProvisioningMessagesDelegate(string message, ProvisioningMessageType messageType);

    /// <summary>
    /// Handles methods for applying provisioning templates
    /// </summary>
    public partial class ProvisioningTemplateApplyingInformation
    {
        private Handlers handlersToProcess = Handlers.All;
        private List<ExtensibilityHandler> extensibilityHandlers = new List<ExtensibilityHandler>();

        public ProvisioningProgressDelegate ProgressDelegate { get; set; }
        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        /// <summary>
        /// If true then persists template information
        /// </summary>
		public bool PersistTemplateInfo { get; set; } = true;
		/// <summary>
		/// If true, system propertybag entries that start with _, vti_, dlc_ etc. will be overwritten if overwrite = true on the PropertyBagEntry. If not true those keys will be skipped, regardless of the overwrite property of the entry.
		/// </summary>
		public bool OverwriteSystemPropertyBagValues { get; set; }

        /// <summary>
        /// If true, existing navigation nodes of the site (where applicable) will be cleared out before applying the navigation nodes from the template (if any). This setting will override any settings made in the template.
        /// </summary>
        public bool ClearNavigation { get; set; }
        /// <summary>
        /// If true then duplicate id errors when the same importing datarows simply generate a warning don't stop the engine. Reason for this is being able to apply the same template multiple times (Delta handling)
        /// without that failing cause the same record is being added twice
        /// </summary>
        public bool IgnoreDuplicateDataRowErrors { get; set; }

        /// <summary>
        /// If true then any content types in the template will be provisioned to subwebs
        /// </summary>
        public bool ProvisionContentTypesToSubWebs { get; set; }

        /// <summary>
        /// If true then any fields in the template will be provisioned to subwebs
        /// </summary>
        public bool ProvisionFieldsToSubWebs { get; set; }

        /// <summary>
        /// Lists of Handlers to process
        /// </summary>
        public Handlers HandlersToProcess
        {
            get
            {
                return handlersToProcess;
            }
            set
            {
                handlersToProcess = value;
            }
        }

        /// <summary>
        /// List of ExtensibilityHandlers
        /// </summary>
        public List<ExtensibilityHandler> ExtensibilityHandlers
        {
            get
            {
                return extensibilityHandlers;
            }

            set
            {
                extensibilityHandlers = value;
            }
        }
    }
}
