using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal delegate bool ShouldProvisionSequenceTest(Tenant web, Model.Provisioning template);

    internal abstract class ObjectSequenceHandlerBase
    {
        internal bool? _willExtract;
        internal bool? _willProvision;

        private bool _reportProgress = true;
        public abstract string Name { get; }

        public bool ReportProgress
        {
            get { return _reportProgress; }
            set { _reportProgress = value; }
        }

        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        public abstract bool WillProvision(Tenant tenant, Model.Provisioning sequenceTemplate, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract bool WillExtract(Tenant tenant, Model.Provisioning sequenceTemplate, ProvisioningTemplateCreationInformation creationInfo);

        public abstract TokenParser ProvisionObjects(Tenant tenant, Model.Provisioning sequenceTemplate, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract ProvisioningTemplate ExtractObjects(Tenant tenant, Model.Provisioning sequenceTemplate, ProvisioningTemplateCreationInformation creationInfo);

        internal void WriteMessage(string message, ProvisioningMessageType messageType)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate(message, messageType);
            }
        }
    }
}
