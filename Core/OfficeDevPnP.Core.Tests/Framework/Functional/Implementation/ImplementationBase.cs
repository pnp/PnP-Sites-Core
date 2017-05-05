using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Implementation
{
    internal abstract class ImplementationBase
    {
        #region Apply template and read the "result"
        public TestProvisioningTemplateResult TestProvisioningTemplate(ClientContext cc, string templateName, Handlers handlersToProcess = Handlers.All, ProvisioningTemplateApplyingInformation ptai = null, ProvisioningTemplateCreationInformation ptci = null)
        {
            try
            {
                // Read the template from XML and apply it
                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(string.Format(@"{0}\..\..\Framework\Functional", AppDomain.CurrentDomain.BaseDirectory), "Templates");
                ProvisioningTemplate sourceTemplate = provider.GetTemplate(templateName);

                if (ptai == null)
                {
                    ptai = new ProvisioningTemplateApplyingInformation();
                    ptai.HandlersToProcess = handlersToProcess;
                }

                if (ptai.ProgressDelegate == null)
                {
                    ptai.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine("Applying template - {0}/{1} - {2}", progress, total, message);
                    };
                }

                sourceTemplate.Connector = provider.Connector;

                TokenParser sourceTokenParser = new TokenParser(cc.Web, sourceTemplate);

                cc.Web.ApplyProvisioningTemplate(sourceTemplate, ptai);

                // Read the site we applied the template to 
                if (ptci == null)
                {
                    ptci = new ProvisioningTemplateCreationInformation(cc.Web);
                    ptci.HandlersToProcess = handlersToProcess;
                }

                if (ptci.ProgressDelegate == null)
                {
                    ptci.ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine("Getting template - {0}/{1} - {2}", progress, total, message);
                    };
                }

                ProvisioningTemplate targetTemplate = cc.Web.GetProvisioningTemplate(ptci);

                return new TestProvisioningTemplateResult()
                {
                    SourceTemplate = sourceTemplate,
                    SourceTokenParser = sourceTokenParser,
                    TargetTemplate = targetTemplate,
                    TargetTokenParser = new TokenParser(cc.Web, targetTemplate),
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString(cc));
                throw;
            }
        }
        #endregion

    }
}
