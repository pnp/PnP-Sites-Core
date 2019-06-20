#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal delegate bool ShouldProvisionSequenceTest(Tenant web, Model.ProvisioningHierarchy hierarchy);

    internal abstract class ObjectHierarchyHandlerBase
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

        public abstract bool WillProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract bool WillExtract(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateCreationInformation creationInfo);

        public abstract TokenParser ProvisionObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation);

        public abstract ProvisioningHierarchy ExtractObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, ProvisioningTemplateCreationInformation creationInfo);

        internal void WriteMessage(string message, ProvisioningMessageType messageType)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate(message, messageType);
            }
        }
    }
}
#endif