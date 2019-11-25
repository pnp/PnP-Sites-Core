#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Configuration;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using OfficeDevPnP.Core.Sites;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectHierarchySequenceSites : ObjectHierarchyHandlerBase
    {
        private List<TokenDefinition> _additionalTokens = new List<TokenDefinition>();
        public override string Name => "Sequences";

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ExtractConfiguration configuration)
        {
            ProvisioningHierarchy tenantTemplate = new ProvisioningHierarchy();
            List<string> siteCollectionUrls = configuration.Tenant.Sequence.SiteUrls;

            List<string> connectedSiteUrls = new List<string>();

            foreach (var siteCollectionUrl in siteCollectionUrls)
            {
                using (var siteContext = tenant.Context.Clone(siteCollectionUrl))
                {
                    if (configuration.Tenant.Sequence.IncludeJoinedSites && siteContext.Site.EnsureProperty(s => s.IsHubSite))
                    {
                        foreach (var hubsiteChildUrl in tenant.GetHubSiteChildUrls(siteContext.Site.EnsureProperty(s => s.Id)))
                        {
                            if (!connectedSiteUrls.Contains(hubsiteChildUrl) && !siteCollectionUrl.Contains(hubsiteChildUrl))
                            {
                                connectedSiteUrls.Add(hubsiteChildUrl);
                            }
                        }
                    }
                }
            }
            siteCollectionUrls.AddRange(connectedSiteUrls);

            ProvisioningSequence provisioningSequence = new ProvisioningSequence();
            provisioningSequence.ID = "TENANTSEQUENCE";
            foreach (var siteCollectionUrl in siteCollectionUrls)
            {
                var siteProperties = tenant.GetSitePropertiesByUrl(siteCollectionUrl, true);

                tenant.Context.Load(siteProperties);
                tenant.Context.ExecuteQueryRetry();
                Model.SiteCollection siteCollection = null;
                using (var siteContext = tenant.Context.Clone(siteCollectionUrl))
                {
                    siteContext.Site.EnsureProperties(s => s.Id, s => s.ShareByEmailEnabled);
                    var templateGuid = siteContext.Site.Id.ToString("N");
                    switch (siteProperties.Template)
                    {
                        case "SITEPAGEPUBLISHING#0":
                            {
                                siteCollection = new CommunicationSiteCollection();

                                siteCollection.IsHubSite = siteProperties.IsHubSite;
                                if (siteProperties.IsHubSite)
                                {
                                    var hubsiteProperties = tenant.GetHubSitePropertiesByUrl(siteCollectionUrl);
                                    tenant.Context.Load(hubsiteProperties);
                                    tenant.Context.ExecuteQueryRetry();
                                    siteCollection.HubSiteLogoUrl = hubsiteProperties.LogoUrl;
                                    siteCollection.HubSiteTitle = hubsiteProperties.Title;
                                }
                                siteCollection.Description = siteProperties.Description;
                                ((CommunicationSiteCollection)siteCollection).Language = (int)siteProperties.Lcid;
                                ((CommunicationSiteCollection)siteCollection).Owner = siteProperties.OwnerEmail;
                                ((CommunicationSiteCollection)siteCollection).AllowFileSharingForGuestUsers = siteContext.Site.ShareByEmailEnabled;
                                tenantTemplate.Parameters.Add($"SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_URL", siteProperties.Url);
                                ((CommunicationSiteCollection)siteCollection).Url = $"{{parameter:SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_URL}}";
                                tenantTemplate.Parameters.Add($"SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_TITLE", siteProperties.Title);
                                siteCollection.Title = $"{{parameter:SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_TITLE}}";
                                break;
                            }
                        case "GROUP#0":
                            {
                                siteCollection = new TeamSiteCollection();
                                siteCollection.IsHubSite = siteProperties.IsHubSite;
                                if (siteProperties.IsHubSite)
                                {
                                    var hubsiteProperties = tenant.GetHubSitePropertiesByUrl(siteCollectionUrl);
                                    tenant.Context.Load(hubsiteProperties);
                                    tenant.Context.ExecuteQueryRetry();
                                    siteCollection.HubSiteLogoUrl = hubsiteProperties.LogoUrl;
                                    siteCollection.HubSiteTitle = hubsiteProperties.Title;
                                }
                                siteCollection.Description = siteProperties.Description;

                                tenantTemplate.Parameters.Add($"SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_ALIAS", siteProperties.Url.Substring(siteProperties.Url.LastIndexOf("/")));
                                ((TeamSiteCollection)siteCollection).Alias = $"{{parameter:SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_ALIAS}}";
                                ((TeamSiteCollection)siteCollection).DisplayName = siteProperties.Title;
                                ((TeamSiteCollection)siteCollection).HideTeamify = Core.Sites.SiteCollection.IsTeamifyPromptHiddenAsync(siteContext).GetAwaiter().GetResult();

                                tenantTemplate.Parameters.Add($"SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_TITLE", siteProperties.Title);
                                siteCollection.Title = $"{{parameter:SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_TITLE}}";
                                break;
                            }
                        case "STS#3":
                            {
                                siteCollection = new TeamNoGroupSiteCollection();
                                siteCollection.IsHubSite = siteProperties.IsHubSite;
                                if (siteProperties.IsHubSite)
                                {
                                    var hubsiteProperties = tenant.GetHubSitePropertiesByUrl(siteCollectionUrl);
                                    tenant.Context.Load(hubsiteProperties);
                                    tenant.Context.ExecuteQueryRetry();
                                    siteCollection.HubSiteLogoUrl = hubsiteProperties.LogoUrl;
                                    siteCollection.HubSiteTitle = hubsiteProperties.Title;
                                }
                                siteCollection.Description = siteProperties.Description;
                                ((TeamNoGroupSiteCollection)siteCollection).Language = (int)siteProperties.Lcid;
                                ((TeamNoGroupSiteCollection)siteCollection).Owner = siteProperties.OwnerEmail;
                                ((TeamNoGroupSiteCollection)siteCollection).TimeZoneId = siteProperties.TimeZoneId;
                                tenantTemplate.Parameters.Add($"SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_URL", siteProperties.Url);
                                ((TeamNoGroupSiteCollection)siteCollection).Url = $"{{parameter:SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_URL}}";
                                tenantTemplate.Parameters.Add($"SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_TITLE", siteProperties.Title);
                                siteCollection.Title = $"{{parameter:SITECOLLECTION_{siteContext.Site.Id.ToString("N")}_TITLE}}";
                                break;
                            }
                    }
                    var siteTemplateCreationInfo = new ProvisioningTemplateCreationInformation(siteContext.Web);

                    // Retrieve the template for the site
                    if (configuration != null)
                    {
                        siteTemplateCreationInfo = configuration.ToCreationInformation(siteContext.Web);
                    }
                    var siteTemplate = siteContext.Web.GetProvisioningTemplate(siteTemplateCreationInfo);
                    siteTemplate.Id = $"TEMPLATE-{templateGuid}";
                    if (siteProperties.HubSiteId != null && siteProperties.HubSiteId != Guid.Empty && siteProperties.HubSiteId != siteContext.Site.Id && siteTemplate.WebSettings != null)
                    {
                        siteTemplate.WebSettings.HubSiteUrl = $"{{parameter:SITECOLLECTION_{siteProperties.HubSiteId.ToString("N")}_URL}}";
                    }
                    tenantTemplate.Templates.Add(siteTemplate);

                    siteCollection.Templates.Add(siteTemplate.Id);

                    if (siteProperties.WebsCount > 1 && configuration.Tenant.Sequence.IncludeSubsites)
                    {
                        var webs = siteContext.Web.EnsureProperty(w => w.Webs);
                        int currentDepth = 1;
                        foreach (var subweb in webs)
                        {
                            siteCollection.Sites.Add(ParseSubsiteSequences(subweb, ref tenantTemplate, configuration, currentDepth, configuration.Tenant.Sequence.MaxSubsiteDepth));
                        }
                    }
                    provisioningSequence.SiteCollections.Add(siteCollection);
                }
            }

            tenantTemplate.Sequences.Add(provisioningSequence);

            PnPProvisioningContext.Current.ParsedSiteUrls.Clear();
            PnPProvisioningContext.Current.ParsedSiteUrls.AddRange(siteCollectionUrls);

            return tenantTemplate;
        }

        private SubSite ParseSubsiteSequences(Web subweb, ref ProvisioningHierarchy tenantTemplate, ExtractConfiguration configuration, int currentDepth, int maxDepth)
        {
            subweb.EnsureProperties(sw => sw.Url, sw => sw.Title, sw => sw.QuickLaunchEnabled, sw => sw.Description, sw => sw.Language, sw => sw.RegionalSettings.TimeZone, sw => sw.Webs, sw => sw.HasUniqueRoleAssignments);

            var subwebTemplate = subweb.GetProvisioningTemplate(configuration.ToCreationInformation(subweb));
            var uniqueid = subweb.Id.ToString("N");
            subwebTemplate.Id = $"TEMPLATE-{uniqueid}";

            tenantTemplate.Templates.Add(subwebTemplate);

            tenantTemplate.Parameters.Add($"SUBSITE_{uniqueid}_URL", subweb.Url.Substring(subweb.Url.LastIndexOf("/")));
            tenantTemplate.Parameters.Add($"SUBSITE_{uniqueid}_TITLE", subweb.Title);
            var subSiteCollection = new TeamNoGroupSubSite()
            {
                Url = $"{{parameter:SUBSITE_{uniqueid}_URL}}",
                Title = $"{{parameter:SUBSITE_{uniqueid}_TITLE}}",
                QuickLaunchEnabled = subweb.QuickLaunchEnabled,
                Description = subweb.Description,
                Language = (int)subweb.Language,
                TimeZoneId = subweb.RegionalSettings.TimeZone.Id,
                UseSamePermissionsAsParentSite = !subweb.HasUniqueRoleAssignments,
                Templates = { subwebTemplate.Id }
            };
            bool traverse = true;
            if (maxDepth != 0)
            {
                currentDepth++;
                traverse = currentDepth <= maxDepth;
            }
            if (traverse && subweb.Webs.AreItemsAvailable)
            {
                currentDepth++;
                foreach (var subsubweb in subweb.Webs)
                {
                    subSiteCollection.Sites.Add(ParseSubsiteSequences(subsubweb, ref tenantTemplate, configuration, currentDepth, maxDepth));
                }
            }
            return subSiteCollection;
        }

        public override TokenParser ProvisionObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, TokenParser tokenParser, ApplyConfiguration configuration)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {
                bool nowait = false;
                if (configuration != null)
                {
                    nowait = configuration.Tenant.DoNotWaitForSitesToBeFullyCreated;
                }
                var sequence = hierarchy.Sequences.FirstOrDefault(s => s.ID == sequenceId);
                if (sequence != null)
                {
                    var siteUrls = new Dictionary<Guid, string>();

                    TokenParser siteTokenParser = null;


                    foreach (var sitecollection in sequence.SiteCollections)
                    {
                        ClientContext siteContext = null;

                        switch (sitecollection)
                        {
                            case TeamSiteCollection t:
                                {
                                    TeamSiteCollectionCreationInformation siteInfo = new TeamSiteCollectionCreationInformation()
                                    {
                                        Alias = tokenParser.ParseString(t.Alias),
                                        DisplayName = tokenParser.ParseString(t.Title),
                                        Description = tokenParser.ParseString(t.Description),
                                        Classification = tokenParser.ParseString(t.Classification),
                                        IsPublic = t.IsPublic,
                                        Lcid = (uint)t.Language
                                    };

                                    var groupSiteInfo = Sites.SiteCollection.GetGroupInfoAsync(tenant.Context as ClientContext, siteInfo.Alias).GetAwaiter().GetResult();
                                    if (groupSiteInfo == null)
                                    {
                                        WriteMessage($"Creating Team Site {siteInfo.Alias}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.Create(tenant.Context as ClientContext, siteInfo, configuration.Tenant.DelayAfterModernSiteCreation, noWait: nowait);
                                    }
                                    else
                                    {
                                        if (groupSiteInfo.ContainsKey("siteUrl"))
                                        {
                                            WriteMessage($"Using existing Team Site {siteInfo.Alias}", ProvisioningMessageType.Progress);
                                            siteContext = (tenant.Context as ClientContext).Clone(groupSiteInfo["siteUrl"], configuration.AccessTokens);
                                        }
                                    }
                                    if (t.IsHubSite)
                                    {
                                        siteContext.Load(siteContext.Site, s => s.Id);
                                        siteContext.ExecuteQueryRetry();
                                        RegisterAsHubSite(tenant, siteContext.Url, siteContext.Site.Id, t.HubSiteLogoUrl, t.HubSiteTitle, tokenParser);
                                    }
                                    if (!string.IsNullOrEmpty(t.Theme))
                                    {
                                        var parsedTheme = tokenParser.ParseString(t.Theme);
                                        tenant.SetWebTheme(parsedTheme, siteContext.Url);
                                        tenant.Context.ExecuteQueryRetry();
                                    }
                                    if (t.Teamify)
                                    {
                                        try
                                        {
                                            WriteMessage($"Teamifying the O365 group connected site at URL - {siteContext.Url}", ProvisioningMessageType.Progress);
                                            siteContext.TeamifyAsync().GetAwaiter().GetResult();
                                        }
                                        catch (Exception ex)
                                        {
                                            WriteMessage($"Teamifying site at URL - {siteContext.Url} failed due to an exception:- {ex.Message}", ProvisioningMessageType.Warning);
                                        }
                                    }
                                    if (t.HideTeamify)
                                    {
                                        try
                                        {
                                            WriteMessage($"Teamify prompt is now hidden for site at URL - {siteContext.Url}", ProvisioningMessageType.Progress);
                                            siteContext.HideTeamifyPrompt().GetAwaiter().GetResult();
                                        }
                                        catch (Exception ex)
                                        {
                                            WriteMessage($"Teamify prompt couldn't be hidden for site at URL - {siteContext.Url} due to an exception:- {ex.Message}", ProvisioningMessageType.Warning);
                                        }
                                    }
                                    siteUrls.Add(t.Id, siteContext.Url);
                                    if (!string.IsNullOrEmpty(t.ProvisioningId))
                                    {
                                        _additionalTokens.Add(new SequenceSiteUrlUrlToken(null, t.ProvisioningId, siteContext.Url));
                                        siteContext.Web.EnsureProperty(w => w.Id);
                                        _additionalTokens.Add(new SequenceSiteIdToken(null, t.ProvisioningId, siteContext.Web.Id));
                                        siteContext.Site.EnsureProperties(s => s.Id, s => s.GroupId);
                                        _additionalTokens.Add(new SequenceSiteCollectionIdToken(null, t.ProvisioningId, siteContext.Site.Id));
                                        _additionalTokens.Add(new SequenceSiteGroupIdToken(null, t.ProvisioningId, siteContext.Site.GroupId));
                                    }
                                    break;
                                }
                            case CommunicationSiteCollection c:
                                {
                                    var siteUrl = tokenParser.ParseString(c.Url);
                                    if (!siteUrl.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        var rootSiteUrl = tenant.GetRootSiteUrl();
                                        tenant.Context.ExecuteQueryRetry();
                                        siteUrl = UrlUtility.Combine(rootSiteUrl.Value, siteUrl);
                                    }
                                    CommunicationSiteCollectionCreationInformation siteInfo = new CommunicationSiteCollectionCreationInformation()
                                    {
                                        ShareByEmailEnabled = c.AllowFileSharingForGuestUsers,
                                        Classification = tokenParser.ParseString(c.Classification),
                                        Description = tokenParser.ParseString(c.Description),
                                        Lcid = (uint)c.Language,
                                        Owner = tokenParser.ParseString(c.Owner),
                                        Title = tokenParser.ParseString(c.Title),
                                        Url = siteUrl
                                    };
                                    if (Guid.TryParse(c.SiteDesign, out Guid siteDesignId))
                                    {
                                        siteInfo.SiteDesignId = siteDesignId;
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(c.SiteDesign))
                                        {
                                            siteInfo.SiteDesign = (CommunicationSiteDesign)Enum.Parse(typeof(CommunicationSiteDesign), c.SiteDesign);
                                        }
                                        else
                                        {
                                            siteInfo.SiteDesign = CommunicationSiteDesign.Showcase;
                                        }
                                    }
                                    // check if site exists
                                    if (tenant.SiteExistsAnywhere(siteInfo.Url) == SiteExistence.Yes)
                                    {
                                        WriteMessage($"Using existing Communications Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(siteInfo.Url, configuration.AccessTokens);
                                    }
                                    else if (tenant.SiteExistsAnywhere(siteInfo.Url) == SiteExistence.Recycled)
                                    {
                                        var errorMessage = $"The requested Communications Site at {siteInfo.Url} is in the Recycle Bin and cannot be created";
                                        WriteMessage(errorMessage, ProvisioningMessageType.Error);
                                        throw new RecycledSiteException(errorMessage);
                                    }
                                    else
                                    {
                                        WriteMessage($"Creating Communications Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.Create(tenant.Context as ClientContext, siteInfo, configuration.Tenant.DelayAfterModernSiteCreation, noWait: nowait);
                                    }
                                    if (c.IsHubSite)
                                    {
                                        siteContext.Load(siteContext.Site, s => s.Id);
                                        siteContext.ExecuteQueryRetry();
                                        RegisterAsHubSite(tenant, siteInfo.Url, siteContext.Site.Id, c.HubSiteLogoUrl, c.HubSiteTitle, tokenParser);
                                    }
                                    if (!string.IsNullOrEmpty(c.Theme))
                                    {
                                        var parsedTheme = tokenParser.ParseString(c.Theme);
                                        tenant.SetWebTheme(parsedTheme, siteInfo.Url);
                                        tenant.Context.ExecuteQueryRetry();
                                    }
                                    siteUrls.Add(c.Id, siteInfo.Url);
                                    if (!string.IsNullOrEmpty(c.ProvisioningId))
                                    {
                                        _additionalTokens.Add(new SequenceSiteUrlUrlToken(null, c.ProvisioningId, siteInfo.Url));
                                        siteContext.Web.EnsureProperty(w => w.Id);
                                        _additionalTokens.Add(new SequenceSiteIdToken(null, c.ProvisioningId, siteContext.Web.Id));
                                        siteContext.Site.EnsureProperties(s => s.Id, s => s.GroupId);
                                        _additionalTokens.Add(new SequenceSiteCollectionIdToken(null, c.ProvisioningId, siteContext.Site.Id));
                                        _additionalTokens.Add(new SequenceSiteGroupIdToken(null, c.ProvisioningId, siteContext.Site.GroupId));
                                    }
                                    break;
                                }
                            case TeamNoGroupSiteCollection t:
                                {
                                    var siteUrl = tokenParser.ParseString(t.Url);
                                    TeamNoGroupSiteCollectionCreationInformation siteInfo = new TeamNoGroupSiteCollectionCreationInformation()
                                    {
                                        Lcid = (uint)t.Language,
                                        Url = siteUrl,
                                        Title = tokenParser.ParseString(t.Title),
                                        Description = tokenParser.ParseString(t.Description),
                                        Owner = tokenParser.ParseString(t.Owner)
                                    };
                                    if (tenant.SiteExistsAnywhere(siteUrl) == SiteExistence.Yes)
                                    {
                                        WriteMessage($"Using existing Team Site at {siteUrl}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(siteUrl, configuration.AccessTokens);
                                    }
                                    else if (tenant.SiteExistsAnywhere(siteUrl) == SiteExistence.Recycled)
                                    {
                                        var errorMessage = $"The requested Team Site at {siteUrl} is in the Recycle Bin and cannot be created";
                                        WriteMessage(errorMessage, ProvisioningMessageType.Error);
                                        throw new RecycledSiteException(errorMessage);
                                    }
                                    else
                                    {
                                        WriteMessage($"Creating Team Site with no Office 365 group at {siteUrl}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.Create(tenant.Context as ClientContext, siteInfo, configuration.Tenant.DelayAfterModernSiteCreation, noWait: nowait);
                                    }
                                    if (t.IsHubSite)
                                    {
                                        siteContext.Load(siteContext.Site, s => s.Id);
                                        siteContext.ExecuteQueryRetry();
                                        RegisterAsHubSite(tenant, siteContext.Url, siteContext.Site.Id, t.HubSiteLogoUrl, t.HubSiteTitle, tokenParser);
                                    }
                                    if (!string.IsNullOrEmpty(t.Theme))
                                    {
                                        var parsedTheme = tokenParser.ParseString(t.Theme);
                                        tenant.SetWebTheme(parsedTheme, siteContext.Url);
                                        tenant.Context.ExecuteQueryRetry();
                                    }
                                    siteUrls.Add(t.Id, siteContext.Url);
                                    if (!string.IsNullOrEmpty(t.ProvisioningId))
                                    {
                                        _additionalTokens.Add(new SequenceSiteUrlUrlToken(null, t.ProvisioningId, siteContext.Url));
                                        siteContext.Web.EnsureProperty(w => w.Id);
                                        _additionalTokens.Add(new SequenceSiteIdToken(null, t.ProvisioningId, siteContext.Web.Id));
                                        siteContext.Site.EnsureProperties(s => s.Id, s => s.GroupId);
                                        _additionalTokens.Add(new SequenceSiteCollectionIdToken(null, t.ProvisioningId, siteContext.Site.Id));
                                        _additionalTokens.Add(new SequenceSiteGroupIdToken(null, t.ProvisioningId, siteContext.Site.GroupId));
                                    }
                                    break;
                                }
                        }

                        var web = siteContext.Web;

                        if (siteTokenParser == null)
                        {
                            siteTokenParser = new TokenParser(tenant, hierarchy, configuration.ToApplyingInformation());
                            foreach (var token in _additionalTokens)
                            {
                                siteTokenParser.AddToken(token);
                            }
                        }

                        foreach (var subsite in sitecollection.Sites)
                        {
                            var subSiteObject = (TeamNoGroupSubSite)subsite;
                            web.EnsureProperties(w => w.Webs.IncludeWithDefaultProperties(), w => w.ServerRelativeUrl);
                            siteTokenParser = CreateSubSites(hierarchy, siteTokenParser, sitecollection, siteContext, web, subSiteObject);
                        }

                        siteTokenParser = null;
                    }

                    // System.Threading.Thread.Sleep(TimeSpan.FromMinutes(10));

                    WriteMessage("Applying templates", ProvisioningMessageType.Progress);
                    var currentSite = "";

                    var provisioningTemplateApplyingInformation = configuration.ToApplyingInformation();
                    provisioningTemplateApplyingInformation.ProgressDelegate = (string message, int step, int total) =>
                    {
                        configuration.ProgressDelegate?.Invoke($"{currentSite} : {message}", step, total);
                    };
                    
                    foreach (var sitecollection in sequence.SiteCollections)
                    {
                        currentSite = sitecollection.ProvisioningId != null ? sitecollection.ProvisioningId : sitecollection.Title;

                        siteUrls.TryGetValue(sitecollection.Id, out string siteUrl);
                        if (siteUrl != null)
                        {
                            using (var clonedContext = tenant.Context.Clone(siteUrl, configuration.AccessTokens))
                            {
                                var web = clonedContext.Web;
                                foreach (var templateRef in sitecollection.Templates)
                                {
                                    var provisioningTemplate = hierarchy.Templates.FirstOrDefault(t => t.Id == templateRef);
                                    if (provisioningTemplate != null)
                                    {
                                        provisioningTemplate.Connector = hierarchy.Connector;
                                        //if (siteTokenParser == null)
                                        //{
                                        siteTokenParser = new TokenParser(web, provisioningTemplate, configuration.ToApplyingInformation());
                                        foreach (var token in _additionalTokens)
                                        {
                                            siteTokenParser.AddToken(token);
                                        }
                                        //}
                                        //else
                                        //{
                                        //    siteTokenParser.Rebase(web, provisioningTemplate);
                                        //}
                                        WriteMessage($"Applying Template", ProvisioningMessageType.Progress);
                                        new SiteToTemplateConversion().ApplyRemoteTemplate(web, provisioningTemplate, provisioningTemplateApplyingInformation, true, siteTokenParser);
                                    }
                                    else
                                    {
                                        WriteMessage($"Referenced template ID {templateRef} not found", ProvisioningMessageType.Error);
                                    }

                                }

                                if (siteTokenParser == null)
                                {
                                    siteTokenParser = new TokenParser(tenant, hierarchy, configuration.ToApplyingInformation());
                                    foreach (var token in _additionalTokens)
                                    {
                                        siteTokenParser.AddToken(token);
                                    }
                                }

                                foreach (var subsite in sitecollection.Sites)
                                {
                                    var subSiteObject = (TeamNoGroupSubSite)subsite;
                                    web.EnsureProperties(w => w.Webs.IncludeWithDefaultProperties(), w => w.ServerRelativeUrl);
                                    siteTokenParser = ApplySubSiteTemplates(hierarchy, siteTokenParser, sitecollection, clonedContext, web, subSiteObject, provisioningTemplateApplyingInformation);
                                }

                                if (sitecollection.IsHubSite)
                                {
                                    RESTUtilities.ExecuteGet(web, "/_api/web/hubsitedata(true)").GetAwaiter().GetResult();
                                }

                            }

                        }
                    }
                }
                return tokenParser;
            }
        }

        private static void RegisterAsHubSite(Tenant tenant, string siteUrl, Guid siteId, string logoUrl, string hubsiteTitle, TokenParser parser)
        {
            siteUrl = parser.ParseString(siteUrl);
            var hubSiteProperties = tenant.GetHubSitePropertiesByUrl(siteUrl);
            tenant.Context.Load<HubSiteProperties>(hubSiteProperties);
            tenant.Context.ExecuteQueryRetry();
            if (hubSiteProperties.ServerObjectIsNull == true)
            {
                var ci = new HubSiteCreationInformation();
                ci.SiteId = siteId;
                if (!string.IsNullOrEmpty(logoUrl))
                {
                    ci.LogoUrl = parser.ParseString(logoUrl);
                }
                if (!string.IsNullOrEmpty(hubsiteTitle))
                {
                    ci.Title = parser.ParseString(hubsiteTitle);
                }
                tenant.RegisterHubSiteWithCreationInformation(siteUrl, ci);
                //tenant.Context.Load(hubSiteProperties);
                tenant.Context.ExecuteQueryRetry();
            }
            else
            {
                bool isDirty = false;
                if (!string.IsNullOrEmpty(logoUrl))
                {
                    logoUrl = parser.ParseString(logoUrl);
                    hubSiteProperties.LogoUrl = logoUrl;
                    isDirty = true;
                }
                if (!string.IsNullOrEmpty(hubsiteTitle))
                {
                    hubsiteTitle = parser.ParseString(hubsiteTitle);
                    hubSiteProperties.Title = hubsiteTitle;
                    isDirty = true;
                }
                if (isDirty)
                {
                    hubSiteProperties.Update();
                    tenant.Context.ExecuteQueryRetry();
                }
            }
        }

        private TokenParser CreateSubSites(ProvisioningHierarchy hierarchy, TokenParser tokenParser, Model.SiteCollection sitecollection, ClientContext siteContext, Web web, TeamNoGroupSubSite subSiteObject)
        {
            var url = tokenParser.ParseString(subSiteObject.Url);

            var subweb = web.Webs.FirstOrDefault(t => t.ServerRelativeUrl.Equals(UrlUtility.Combine(web.ServerRelativeUrl, "/", url.Trim(new char[] { '/' }))));
            if (subweb == null)
            {
                subweb = web.Webs.Add(new WebCreationInformation()
                {
                    Language = subSiteObject.Language,
                    Url = url,
                    Description = tokenParser.ParseString(subSiteObject.Description),
                    Title = tokenParser.ParseString(subSiteObject.Title),
                    UseSamePermissionsAsParentSite = subSiteObject.UseSamePermissionsAsParentSite,
                    WebTemplate = "STS#3"
                });
                WriteMessage($"Creating Sub Site with no Office 365 group at {url}", ProvisioningMessageType.Progress);
                siteContext.Load(subweb);
                siteContext.ExecuteQueryRetry();
            }
            else
            {
                WriteMessage($"Using existing Sub Site with no Office 365 group at {url}", ProvisioningMessageType.Progress);
            }

            if (subSiteObject.Sites.Any())
            {
                foreach (var subsubSite in subSiteObject.Sites)
                {
                    var subsubSiteObject = (TeamNoGroupSubSite)subsubSite;
                    tokenParser = CreateSubSites(hierarchy, tokenParser, sitecollection, siteContext, subweb, subsubSiteObject);
                }
            }

            return tokenParser;
        }

        private TokenParser ApplySubSiteTemplates(ProvisioningHierarchy hierarchy, TokenParser tokenParser, Model.SiteCollection sitecollection, ClientContext siteContext, Web web, TeamNoGroupSubSite subSiteObject, ProvisioningTemplateApplyingInformation provisioningTemplateApplyingInformation)
        {
            var url = tokenParser.ParseString(subSiteObject.Url);

            var subweb = web.Webs.FirstOrDefault(t => t.ServerRelativeUrl.Equals(UrlUtility.Combine(web.ServerRelativeUrl, "/", url.Trim(new char[] { '/' }))));

            foreach (var templateRef in subSiteObject.Templates)
            {
                var provisioningTemplate = hierarchy.Templates.FirstOrDefault(t => t.Id == templateRef);
                if (provisioningTemplate != null)
                {
                    provisioningTemplate.Connector = hierarchy.Connector;
                    if (tokenParser == null)
                    {
                        tokenParser = new TokenParser(subweb, provisioningTemplate);
                    }
                    else
                    {
                        tokenParser.Rebase(subweb, provisioningTemplate, provisioningTemplateApplyingInformation);
                    }
                    new SiteToTemplateConversion().ApplyRemoteTemplate(subweb, provisioningTemplate, provisioningTemplateApplyingInformation, true, tokenParser);
                }
                else
                {
                    WriteMessage($"Referenced template ID {templateRef} not found", ProvisioningMessageType.Error);
                }
            }

            if (subSiteObject.Sites.Any())
            {
                foreach (var subsubSite in subSiteObject.Sites)
                {
                    var subsubSiteObject = (TeamNoGroupSubSite)subsubSite;
                    tokenParser = ApplySubSiteTemplates(hierarchy, tokenParser, sitecollection, siteContext, subweb, subsubSiteObject, provisioningTemplateApplyingInformation);
                }
            }

            return tokenParser;
        }


        public override bool WillExtract(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ExtractConfiguration creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration)
        {
            return hierarchy.Sequences.Count > 0;
        }
    }
}
#endif