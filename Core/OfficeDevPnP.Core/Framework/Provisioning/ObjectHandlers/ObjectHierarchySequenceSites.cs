#if !ONPREMISES
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
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

        public override ProvisioningHierarchy ExtractObjects(Tenant tenant, ProvisioningHierarchy hierarchy, ProvisioningTemplateCreationInformation creationInfo)
        {
            throw new NotImplementedException();
        }

        public override TokenParser ProvisionObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, TokenParser tokenParser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {
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
                                        IsPublic = t.IsPublic
                                    };

                                    var groupSiteInfo = Sites.SiteCollection.GetGroupInfoAsync(tenant.Context as ClientContext, siteInfo.Alias).GetAwaiter().GetResult();
                                    if (groupSiteInfo == null)
                                    {
                                        WriteMessage($"Creating Team Site {siteInfo.Alias}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.Create(tenant.Context as ClientContext, siteInfo, applyingInformation.DelayAfterModernSiteCreation);
                                    }
                                    else
                                    {
                                        if (groupSiteInfo.ContainsKey("siteUrl"))
                                        {
                                            WriteMessage($"Using existing Team Site {siteInfo.Alias}", ProvisioningMessageType.Progress);
                                            siteContext = (tenant.Context as ClientContext).Clone(groupSiteInfo["siteUrl"], applyingInformation.AccessTokens);
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
                                    if (tenant.SiteExists(siteInfo.Url))
                                    {
                                        WriteMessage($"Using existing Communications Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(siteInfo.Url, applyingInformation.AccessTokens);
                                    }
                                    else
                                    {
                                        WriteMessage($"Creating Communications Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.Create(tenant.Context as ClientContext, siteInfo, applyingInformation.DelayAfterModernSiteCreation);
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
                                    if (tenant.SiteExists(siteUrl))
                                    {
                                        WriteMessage($"Using existing Team Site at {siteUrl}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(siteUrl, applyingInformation.AccessTokens);
                                    }
                                    else
                                    {
                                        WriteMessage($"Creating Team Site with no Office 365 group at {siteUrl}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.Create(tenant.Context as ClientContext, siteInfo, applyingInformation.DelayAfterModernSiteCreation);
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
                            siteTokenParser = new TokenParser(tenant, hierarchy, applyingInformation);
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

                    var provisioningTemplateApplyingInformation = new ProvisioningTemplateApplyingInformation();
                    provisioningTemplateApplyingInformation.AccessTokens = applyingInformation.AccessTokens;
                    provisioningTemplateApplyingInformation.MessagesDelegate = applyingInformation.MessagesDelegate;
                    provisioningTemplateApplyingInformation.ProgressDelegate = (string message, int step, int total) =>
                    {
                        applyingInformation.ProgressDelegate?.Invoke($"{currentSite} : {message}", step, total);
                    };
                    provisioningTemplateApplyingInformation.SiteProvisionedDelegate = applyingInformation.SiteProvisionedDelegate;

                    foreach (var sitecollection in sequence.SiteCollections)
                    {
                        currentSite = sitecollection.ProvisioningId != null ? sitecollection.ProvisioningId : sitecollection.Title;

                        siteUrls.TryGetValue(sitecollection.Id, out string siteUrl);
                        if (siteUrl != null)
                        {
                            using (var clonedContext = tenant.Context.Clone(siteUrl, applyingInformation.AccessTokens))
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
                                        siteTokenParser = new TokenParser(web, provisioningTemplate, applyingInformation);
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
                                    siteTokenParser = new TokenParser(tenant, hierarchy, applyingInformation);
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


        public override bool WillExtract(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateCreationInformation creationInfo)
        {
            throw new NotImplementedException();
        }

        public override bool WillProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return hierarchy.Sequences.Count > 0;
        }
    }
}
#endif