using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class SiteToTemplateConversion
    {
        /// <summary>
        /// Actual implementation of extracting configuration from existing site.
        /// </summary>
        /// <param name="web"></param>
        /// <param name="creationInfo"></param>
        /// <returns></returns>
        internal ProvisioningTemplate GetRemoteTemplate(Web web, ProvisioningTemplateCreationInformation creationInfo)
        {

            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Extraction))
            {

                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                if (creationInfo != null)
                {
                    if (creationInfo.BaseTemplate != null)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_Base_template_available___0_, creationInfo.BaseTemplate.Id);
                    }
                    progressDelegate = creationInfo.ProgressDelegate;
                    if (creationInfo.ProgressDelegate != null)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_ProgressDelegate_registered);
                    }
                    messagesDelegate = creationInfo.MessagesDelegate;
                    if (creationInfo.MessagesDelegate != null)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_MessagesDelegate_registered);
                    }
                    if (creationInfo.IncludeAllTermGroups)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_IncludeAllTermGroups_is_set_to_true);
                    }
                    if (creationInfo.IncludeSiteCollectionTermGroup)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_IncludeSiteCollectionTermGroup_is_set_to_true);
                    }
                    if (creationInfo.PersistComposedLookFiles)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_PersistComposedLookFiles_is_set_to_true);
                    }
                }

                // Create empty object
                ProvisioningTemplate template = new ProvisioningTemplate();

                // Hookup connector, is handy when the generated template object is used to apply to another site
                template.Connector = creationInfo.FileConnector;

                List<ObjectHandlerBase> objectHandlers = new List<ObjectHandlerBase>();

                objectHandlers.Add(new ObjectRegionalSettings());
                objectHandlers.Add(new ObjectSupportedUILanguages());
                objectHandlers.Add(new ObjectAuditSettings());
                objectHandlers.Add(new ObjectSitePolicy());
                objectHandlers.Add(new ObjectSiteSecurity());
                objectHandlers.Add(new ObjectTermGroups());
                objectHandlers.Add(new ObjectField());
                objectHandlers.Add(new ObjectContentType());
                objectHandlers.Add(new ObjectListInstance());
                objectHandlers.Add(new ObjectCustomActions());
                objectHandlers.Add(new ObjectFeatures());
                objectHandlers.Add(new ObjectComposedLook());
                objectHandlers.Add(new ObjectSearchSettings());
                objectHandlers.Add(new ObjectFiles());
                objectHandlers.Add(new ObjectPages());
                objectHandlers.Add(new ObjectPropertyBagEntry());
                objectHandlers.Add(new ObjectPublishing());
                objectHandlers.Add(new ObjectWorkflows());
                objectHandlers.Add(new ObjectRetrieveTemplateInfo());

                int step = 1;

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillExtract(web, template, creationInfo));

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillExtract(web, template, creationInfo))
                    {
                        if (messagesDelegate != null)
                        {
                            handler.MessagesDelegate = messagesDelegate;
                        }
                        if (handler.ReportProgress && progressDelegate != null)
                        {
                            progressDelegate(handler.Name, step, count);
                            step++;
                        }
                        template = handler.ExtractObjects(web, template, creationInfo);
                    }
                }
                return template;
            }
        }

        /// <summary>
        /// Actual implementation of the apply templates
        /// </summary>
        /// <param name="web"></param>
        /// <param name="template"></param>
        /// <param name="provisioningInfo"></param>
        internal void ApplyRemoteTemplate(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation provisioningInfo)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {
                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                if (provisioningInfo != null)
                {
                    if (provisioningInfo.OverwriteSystemPropertyBagValues == true)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_ApplyRemoteTemplate_OverwriteSystemPropertyBagValues_is_to_true);
                    }
                    progressDelegate = provisioningInfo.ProgressDelegate;
                    if (provisioningInfo.ProgressDelegate != null)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_ProgressDelegate_registered);
                    }
                    messagesDelegate = provisioningInfo.MessagesDelegate;
                    if (provisioningInfo.MessagesDelegate != null)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_MessagesDelegate_registered);
                    }
                }

                List<ObjectHandlerBase> objectHandlers = new List<ObjectHandlerBase>();

                objectHandlers.Add(new ObjectRegionalSettings());
                objectHandlers.Add(new ObjectSupportedUILanguages());
                objectHandlers.Add(new ObjectAuditSettings());
                objectHandlers.Add(new ObjectSitePolicy());
                objectHandlers.Add(new ObjectSiteSecurity());
                objectHandlers.Add(new ObjectFeatures());
                objectHandlers.Add(new ObjectTermGroups());
                objectHandlers.Add(new ObjectField());
                objectHandlers.Add(new ObjectContentType());
                objectHandlers.Add(new ObjectListInstance());
                objectHandlers.Add(new ObjectLookupFields());
                objectHandlers.Add(new ObjectListInstanceDataRows());
                objectHandlers.Add(new ObjectFiles());
                objectHandlers.Add(new ObjectPages());
                objectHandlers.Add(new ObjectCustomActions());
                objectHandlers.Add(new ObjectPublishing());
                objectHandlers.Add(new ObjectComposedLook());
                objectHandlers.Add(new ObjectSearchSettings());
                objectHandlers.Add(new ObjectWorkflows());
                objectHandlers.Add(new ObjectPropertyBagEntry());
                objectHandlers.Add(new ObjectExtensibilityProviders());
                objectHandlers.Add(new ObjectPersistTemplateInfo());

                var tokenParser = new TokenParser(web, template);

                int step = 1;

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(web, template));

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillProvision(web, template))
                    {
                        if (messagesDelegate != null)
                        {
                            handler.MessagesDelegate = messagesDelegate;
                        }
                        if (handler.ReportProgress && progressDelegate != null)
                        {
                            progressDelegate(handler.Name, step, count);
                            step++;
                        }
                        tokenParser = handler.ProvisionObjects(web, template, tokenParser, provisioningInfo);
                    }
                }
            }
        }
    }
}
