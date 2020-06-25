using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal class WorkflowImplementation : ImplementationBase
    {
        internal void Workflows(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                ptci.HandlersToProcess = Handlers.Lists | Handlers.Workflows;
                ptci.FileConnector = new FileSystemConnector(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");

                string xmlFileName = null;
#if SP2013
                // pnp:WorkflowSubscription ParentContentTypeId="" not availiable for comparing
                xmlFileName = "workflows_add_1605.SP2013.xml";
#else
                xmlFileName = "workflows_add_1605.xml";
#endif

                var result = TestProvisioningTemplate(cc, xmlFileName, Handlers.Lists | Handlers.Workflows, null, ptci);
                WorkflowValidator wv = new WorkflowValidator();
                Assert.IsTrue(wv.Validate(result.SourceTemplate.Workflows, result.TargetTemplate.Workflows, result.TargetTokenParser));
            }
        }
    }
}