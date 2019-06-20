using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;

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

#if !ONPREMISES || SP2016 || SP2019
                web.Context.DisableReturnValueCache = true;
#endif

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
                    if (creationInfo.PersistBrandingFiles)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_PersistBrandingFiles_is_set_to_true);
                    }
                }
                else
                {
                    // When no provisioning info was passed then we want to execute all handlers
                    creationInfo = new ProvisioningTemplateCreationInformation(web);
                    creationInfo.HandlersToProcess = Handlers.All;
                }

                // Create empty object
                ProvisioningTemplate template = new ProvisioningTemplate();

                // Hookup connector, is handy when the generated template object is used to apply to another site
                template.Connector = creationInfo.FileConnector;

                List<ObjectHandlerBase> objectHandlers = new List<ObjectHandlerBase>();

                if (creationInfo.HandlersToProcess.HasFlag(Handlers.RegionalSettings)) objectHandlers.Add(new ObjectRegionalSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SupportedUILanguages)) objectHandlers.Add(new ObjectSupportedUILanguages());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.AuditSettings)) objectHandlers.Add(new ObjectAuditSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SitePolicy)) objectHandlers.Add(new ObjectSitePolicy());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteSecurity)) objectHandlers.Add(new ObjectSiteSecurity());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.TermGroups)) objectHandlers.Add(new ObjectTermGroups());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Fields)) objectHandlers.Add(new ObjectField(FieldAndListProvisioningStepHelper.Step.Export));
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes)) objectHandlers.Add(new ObjectContentType(FieldAndListProvisioningStepHelper.Step.Export));
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.Export));
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.CustomActions)) objectHandlers.Add(new ObjectCustomActions());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Features)) objectHandlers.Add(new ObjectFeatures());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ComposedLook)) objectHandlers.Add(new ObjectComposedLook());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SearchSettings)) objectHandlers.Add(new ObjectSearchSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Files)) objectHandlers.Add(new ObjectFiles());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Pages)) objectHandlers.Add(new ObjectPages());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.PageContents)) objectHandlers.Add(new ObjectPageContents());
#if !ONPREMISES
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.PageContents)) objectHandlers.Add(new ObjectClientSidePageContents());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteHeader)) objectHandlers.Add(new ObjectSiteHeaderSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteFooter)) objectHandlers.Add(new ObjectSiteFooterSettings());
#endif
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.PropertyBagEntries)) objectHandlers.Add(new ObjectPropertyBagEntry());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Publishing)) objectHandlers.Add(new ObjectPublishing());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Workflows)) objectHandlers.Add(new ObjectWorkflows());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.WebSettings)) objectHandlers.Add(new ObjectWebSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Theme)) objectHandlers.Add(new ObjectTheme());

                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Navigation)) objectHandlers.Add(new ObjectNavigation());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ImageRenditions)) objectHandlers.Add(new ObjectImageRenditions());
                objectHandlers.Add(new ObjectLocalization()); // Always add this one, check is done in the handler
#if !ONPREMISES
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Tenant)) objectHandlers.Add(new ObjectTenant());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ApplicationLifecycleManagement)) objectHandlers.Add(new ObjectApplicationLifecycleManagement());
#endif
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ExtensibilityProviders)) objectHandlers.Add(new ObjectExtensibilityHandlers());

                objectHandlers.Add(new ObjectRetrieveTemplateInfo());

                int step = 1;

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillExtract(web, template, creationInfo));

                web.EnsureProperty(w => w.Url);

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

                        using (var handlerContext = web.Context.Clone(web.Url))
                        {
                            template = handler.ExtractObjects(handlerContext.Web, template, creationInfo);
                        }
                    }
                }

                return template;
            }
        }

#if !ONPREMISES
        internal void ApplyProvisioningHierarchy(Tenant tenant, OfficeDevPnP.Core.Framework.Provisioning.Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation provisioningInfo)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {
                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                if (provisioningInfo == null)
                {
                    // When no provisioning info was passed then we want to execute all handlers
                    provisioningInfo = new ProvisioningTemplateApplyingInformation();
                    provisioningInfo.HandlersToProcess = Handlers.All;
                }
                else
                {
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
                    if (provisioningInfo.HandlersToProcess == default(Handlers))
                    {
                        provisioningInfo.HandlersToProcess = Handlers.All;
                    }

                }

                List<ObjectHierarchyHandlerBase> objectHandlers = new List<ObjectHierarchyHandlerBase>
                {
                    new ObjectHierarchyTenant(),
                    new ObjectHierarchySequenceTermGroups(),
                    new ObjectHierarchySequenceSites(),
                    new ObjectTeams(),
                    new ObjectAzureActiveDirectory(),
                };

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(tenant, hierarchy, sequenceId, provisioningInfo)) + 1;

                progressDelegate?.Invoke("Initializing engine", 1, count); // handlers + initializing message)

                int step = 2;

                TokenParser sequenceTokenParser = new TokenParser(tenant, hierarchy);
                foreach (var handler in objectHandlers)
                {
                    if (handler.WillProvision(tenant, hierarchy, sequenceId, provisioningInfo))
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
                        sequenceTokenParser = handler.ProvisionObjects(tenant, hierarchy, sequenceId, sequenceTokenParser, provisioningInfo);
                    }
                }
            }
        }
#endif
        /// <summary>
        /// Actual implementation of the apply templates
        /// </summary>
        /// <param name="web"></param>
        /// <param name="template"></param>
        /// <param name="provisioningInfo"></param>
        /// <param name="calledFromHierarchy"></param>
        /// <param name="tokenParser"></param>
        internal void ApplyRemoteTemplate(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation provisioningInfo, bool calledFromHierarchy = false, TokenParser tokenParser = null)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {

#if !ONPREMISES || SP2016 || SP2019
                web.Context.DisableReturnValueCache = true;
#endif

                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                ProvisioningSiteProvisionedDelegate siteProvisionedDelegate = null;
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
                    siteProvisionedDelegate = provisioningInfo.SiteProvisionedDelegate;
                }
                else
                {
                    // When no provisioning info was passed then we want to execute all handlers
                    provisioningInfo = new ProvisioningTemplateApplyingInformation();
                    provisioningInfo.HandlersToProcess = Handlers.All;
                }

                // Check if scope is present and if so, matches the current site. When scope was not set the returned value will be ProvisioningTemplateScope.Undefined
                if (template.Scope == ProvisioningTemplateScope.RootSite)
                {
                    if (web.IsSubSite())
                    {
                        scope.LogError(CoreResources.SiteToTemplateConversion_ScopeOfTemplateDoesNotMatchTarget);
                        throw new Exception(CoreResources.SiteToTemplateConversion_ScopeOfTemplateDoesNotMatchTarget);
                    }
                }
                var currentCultureInfoValue = System.Threading.Thread.CurrentThread.CurrentCulture.LCID;
                if (!string.IsNullOrEmpty(template.TemplateCultureInfo))
                {
                    int cultureInfoValue = System.Threading.Thread.CurrentThread.CurrentCulture.LCID;
                    if (int.TryParse(template.TemplateCultureInfo, out cultureInfoValue))
                    {
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(cultureInfoValue);
                    }
                    else
                    {
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(template.TemplateCultureInfo);
                    }
                }

                // Check if the target site shares the same base template with the template's source site
                var targetSiteTemplateId = web.GetBaseTemplateId();
                if (!String.IsNullOrEmpty(targetSiteTemplateId) && !String.IsNullOrEmpty(template.BaseSiteTemplate))
                {
                    if (!targetSiteTemplateId.Equals(template.BaseSiteTemplate, StringComparison.InvariantCultureIgnoreCase))
                    {
                        var templatesNotMatchingWarning = String.Format(CoreResources.Provisioning_Asymmetric_Base_Templates, template.BaseSiteTemplate, targetSiteTemplateId);
                        scope.LogWarning(templatesNotMatchingWarning);
                        messagesDelegate?.Invoke(templatesNotMatchingWarning, ProvisioningMessageType.Warning);
                    }
                }

                // Always ensure the Url property is loaded. In the tokens we need this and we don't want to call ExecuteQuery as this can
                // impact delta scenarions (calling ExecuteQuery before the planned update is called)
                web.EnsureProperty(w => w.Url);


                List<ObjectHandlerBase> objectHandlers = new List<ObjectHandlerBase>();

                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.RegionalSettings)) objectHandlers.Add(new ObjectRegionalSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SupportedUILanguages)) objectHandlers.Add(new ObjectSupportedUILanguages());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.AuditSettings)) objectHandlers.Add(new ObjectAuditSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SitePolicy)) objectHandlers.Add(new ObjectSitePolicy());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteSecurity)) objectHandlers.Add(new ObjectSiteSecurity());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Features)) objectHandlers.Add(new ObjectFeatures());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.TermGroups)) objectHandlers.Add(new ObjectTermGroups());

                // Process 3 times these providers to handle proper ordering of artefact creation when dealing with lookup fields

                // 1st. create fields, content and list without lookup fields
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Fields) || provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes)) objectHandlers.Add(new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields));

                // 2nd. create lookup fields (which requires lists to be present
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Fields) || provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectField(FieldAndListProvisioningStepHelper.Step.LookupFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes)) objectHandlers.Add(new ObjectContentType(FieldAndListProvisioningStepHelper.Step.LookupFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.LookupFields));

                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Files)) objectHandlers.Add(new ObjectFiles());

                // 3rd. Create remaining objects in lists (views, user custom actions, ...)
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings));

                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Fields) || provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstanceDataRows());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Workflows)) objectHandlers.Add(new ObjectWorkflows());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Pages)) objectHandlers.Add(new ObjectPages());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.PageContents)) objectHandlers.Add(new ObjectPageContents());
#if !ONPREMISES
                if (!calledFromHierarchy && provisioningInfo.HandlersToProcess.HasFlag(Handlers.Tenant)) objectHandlers.Add(new ObjectTenant());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ApplicationLifecycleManagement)) objectHandlers.Add(new ObjectApplicationLifecycleManagement());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Pages)) objectHandlers.Add(new ObjectClientSidePages());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteHeader)) objectHandlers.Add(new ObjectSiteHeaderSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteFooter)) objectHandlers.Add(new ObjectSiteFooterSettings());
#endif
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.CustomActions)) objectHandlers.Add(new ObjectCustomActions());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Publishing)) objectHandlers.Add(new ObjectPublishing());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ComposedLook)) objectHandlers.Add(new ObjectComposedLook());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SearchSettings)) objectHandlers.Add(new ObjectSearchSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.PropertyBagEntries)) objectHandlers.Add(new ObjectPropertyBagEntry());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.WebSettings)) objectHandlers.Add(new ObjectWebSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Theme)) objectHandlers.Add(new ObjectTheme());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Navigation)) objectHandlers.Add(new ObjectNavigation());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ImageRenditions)) objectHandlers.Add(new ObjectImageRenditions());
                objectHandlers.Add(new ObjectLocalization()); // Always add this one, check is done in the handler
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ExtensibilityProviders)) objectHandlers.Add(new ObjectExtensibilityHandlers());

                // Only persist template information in case this flag is set: this will allow the engine to
                // work with lesser permissions
                if (provisioningInfo.PersistTemplateInfo)
                {
                    objectHandlers.Add(new ObjectPersistTemplateInfo());
                }
                var count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(web, template, provisioningInfo)) + 1;

                progressDelegate?.Invoke("Initializing engine", 1, count); // handlers + initializing message)
                if (tokenParser == null)
                {
                    tokenParser = new TokenParser(web, template);
                }
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ExtensibilityProviders))
                {
                    var extensibilityHandler = objectHandlers.OfType<ObjectExtensibilityHandlers>().First();
                    extensibilityHandler.AddExtendedTokens(web, template, tokenParser, provisioningInfo);
                }

                int step = 2;

                // Remove potentially unsupported artifacts

                var cleaner = new NoScriptTemplateCleaner(web);
                if (messagesDelegate != null)
                {
                    cleaner.MessagesDelegate = messagesDelegate;
                }
                template = cleaner.CleanUpBeforeProvisioning(template);

                CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ProvisioningStarted);

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillProvision(web, template, provisioningInfo))
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
                        CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted, handler);
                        tokenParser = handler.ProvisionObjects(web, template, tokenParser, provisioningInfo);
                        CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted, handler);
                    }
                }

                // Notify the completed provisioning of the site
                web.EnsureProperties(w => w.Title, w => w.Url);
                siteProvisionedDelegate?.Invoke(web.Title, web.Url);

                CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ProvisioningCompleted);

                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(currentCultureInfoValue);

            }
        }

        private void CallWebHooks(ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateWebhookKind kind, ObjectHandlerBase objectHandler = null)
        {
            using (var scope = new PnPMonitoredScope("ProvisioningTemplate WebHook Call"))
            {
                if (template.ProvisioningTemplateWebhooks != null && template.ProvisioningTemplateWebhooks.Any())
                {
                    foreach (var webhook in template.ProvisioningTemplateWebhooks.Where(w => w.Kind == kind))
                    {
                        var requestParameters = new Dictionary<String, String>();

                        SimpleTokenParser internalParser = new SimpleTokenParser();
                        foreach (var webhookparam in webhook.Parameters)
                        {
                            requestParameters.Add(webhookparam.Key, parser.ParseString(webhookparam.Value));
                            internalParser.AddToken(new WebhookParameter(webhookparam.Key, requestParameters[webhookparam.Key]));
                        }
                        var url = parser.ParseString(webhook.Url); // parse for template scoped parameters
                        url = internalParser.ParseString(url); // parse for webhook scoped parameters

                        switch (webhook.Method)
                        {
                            case ProvisioningTemplateWebhookMethod.GET:
                                {
                                    if (kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted || kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted)
                                    {
                                        url += $"&__handler={objectHandler.InternalName}"; // add the handler name to the REST request URL
                                        url += $"&__webhookKind={kind.ToString()}"; // add the webhook kind to the REST request URL
                                    }
                                    try
                                    {
                                        using (var client = new HttpClient())
                                        {
                                            if (webhook.Async)
                                            {
                                                client.GetAsync(url);
                                            }
                                            else
                                            {
                                                client.GetAsync(url).GetAwaiter().GetResult();
                                            }
                                        }
                                    }
                                    catch (HttpRequestException ex)
                                    {
                                        scope.LogError(ex, "Error calling provisioning template webhook");
                                    }
                                    break;
                                }
                            case ProvisioningTemplateWebhookMethod.POST:
                                {
                                    requestParameters.Add("__webhookKind", kind.ToString()); // add the webhook kind to the parameters of the request body
                                    if (kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted || kind == ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted)
                                    {
                                        requestParameters.Add("__handler", objectHandler.InternalName); // add the handler name to the parameters of the request body
                                    }
                                    try
                                    {
                                        using (var client = new HttpClient())
                                        {
                                            if (webhook.Async)
                                            {
                                                switch (webhook.BodyFormat)
                                                {
                                                    case ProvisioningTemplateWebhookBodyFormat.Json:
                                                        client.PostAsJsonAsync(url, requestParameters);
                                                        break;
                                                    case ProvisioningTemplateWebhookBodyFormat.Xml:
                                                        client.PostAsXmlAsync(url, requestParameters);
                                                        break;
                                                    case ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded:
                                                        var content = new FormUrlEncodedContent(requestParameters);
                                                        client.PostAsync(url, content);
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (webhook.BodyFormat)
                                                {
                                                    case ProvisioningTemplateWebhookBodyFormat.Json:
                                                        client.PostAsJsonAsync(url, requestParameters).GetAwaiter().GetResult();
                                                        break;
                                                    case ProvisioningTemplateWebhookBodyFormat.Xml:
                                                        client.PostAsXmlAsync(url, requestParameters).GetAwaiter().GetResult();
                                                        break;
                                                    case ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded:
                                                        var content = new FormUrlEncodedContent(requestParameters);
                                                        client.PostAsync(url, content).GetAwaiter().GetResult();
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                    catch (HttpRequestException ex)
                                    {
                                        scope.LogError(ex, "Error calling provisioning template webhook");
                                    }
                                    break;
                                }
                        }
                    }
                }
            }
        }
    }
}
