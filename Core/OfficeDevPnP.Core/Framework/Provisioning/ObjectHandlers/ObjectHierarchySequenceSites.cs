using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Sites;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectHierarchySequenceSites : ObjectHierarchyHandlerBase
    {
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
                                    var groupSiteInfo = Sites.SiteCollection.GetGroupInfo(tenant.Context as ClientContext, siteInfo.Alias).GetAwaiter().GetResult();
                                    if (groupSiteInfo == null)
                                    {
                                        WriteMessage($"Creating Team Site {siteInfo.Alias}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.CreateAsync(tenant.Context as ClientContext, siteInfo).GetAwaiter().GetResult();
                                    }
                                    else
                                    {
                                        if (groupSiteInfo.ContainsKey("siteUrl"))
                                        {
                                            WriteMessage($"Using existing Team Site {siteInfo.Alias}", ProvisioningMessageType.Progress);
                                            siteContext = (tenant.Context as ClientContext).Clone(groupSiteInfo["siteUrl"]);
                                        }
                                    }

                                    break;
                                }
                            case CommunicationSiteCollection c:
                                {

                                    CommunicationSiteCollectionCreationInformation siteInfo = new CommunicationSiteCollectionCreationInformation()
                                    {
                                        AllowFileSharingForGuestUsers = c.AllowFileSharingForGuestUsers,
                                        Classification = tokenParser.ParseString(c.Classification),
                                        Description = tokenParser.ParseString(c.Description),
                                        Lcid = (uint)c.Language,
                                        Owner = tokenParser.ParseString(c.Owner),
                                        Title = tokenParser.ParseString(c.Title),
                                        Url = tokenParser.ParseString(c.Url)
                                    };
                                    if (Guid.TryParse(c.SiteDesign, out Guid siteDesignId))
                                    {
                                        siteInfo.SiteDesignId = siteDesignId;
                                    }
                                    else
                                    {
                                        siteInfo.SiteDesign = (CommunicationSiteDesign)Enum.Parse(typeof(CommunicationSiteDesign), c.SiteDesign);
                                    }
                                    // check if site exists
                                    if (tenant.SiteExists(siteInfo.Url))
                                    {
                                        WriteMessage($"Using existing Communications Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(siteInfo.Url);
                                    }
                                    else
                                    {
                                        WriteMessage($"Creating Communications Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.CreateAsync(tenant.Context as ClientContext, siteInfo).GetAwaiter().GetResult();
                                    }
                                    break;
                                }
                            case TeamNoGroupSiteCollection t:
                                {
                                    SiteEntity siteInfo = new SiteEntity()
                                    {
                                        Lcid = (uint)t.Language,
                                        Template = "STS#3",
                                        TimeZoneId = t.TimeZoneId,
                                        Title = tokenParser.ParseString(t.Title),
                                        Url = tokenParser.ParseString(t.Url),
                                        SiteOwnerLogin = tokenParser.ParseString(t.Owner),
                                    };
                                    WriteMessage($"Creating Team Site with no Office 365 group at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                    if (tenant.SiteExists(t.Url))
                                    {
                                        WriteMessage($"Using existing Team Site at {siteInfo.Url}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(t.Url);
                                    }
                                    else
                                    {
                                        tenant.CreateSiteCollection(siteInfo, false, true);
                                        siteContext = tenant.Context.Clone(t.Url);
                                    }
                                    break;
                                }
                        }
                        var web = siteContext.Web;
                        foreach (var templateRef in sitecollection.Templates)
                        {
                            var provisioningTemplate = hierarchy.Templates.FirstOrDefault(t => t.Id == templateRef);
                            if (provisioningTemplate != null)
                            {
                                if (siteTokenParser == null)
                                {
                                    siteTokenParser = new TokenParser(web, provisioningTemplate);
                                }
                                else
                                {
                                    siteTokenParser.Rebase(web, provisioningTemplate);
                                }
                                WriteMessage($"Applying Template", ProvisioningMessageType.Progress);
                                new SiteToTemplateConversion().ApplyRemoteTemplate(web, provisioningTemplate, new ProvisioningTemplateApplyingInformation(), true, tokenParser);
                            }
                            else
                            {
                                WriteMessage($"Referenced template ID {templateRef} not found", ProvisioningMessageType.Error);
                            }

                        }

                        if (siteTokenParser == null)
                        {
                            siteTokenParser = new TokenParser(tenant, hierarchy);
                        }

                        foreach (var subsite in sitecollection.Sites)
                        {
                            var subSiteObject = (TeamNoGroupSubSite)subsite;
                            web.EnsureProperties(w => w.Webs.IncludeWithDefaultProperties(), w => w.ServerRelativeUrl);
                            siteTokenParser = ParseSubsites(hierarchy, siteTokenParser, sitecollection, siteContext, web, subSiteObject);
                        }
                    }
                }
                return tokenParser;
            }
        }

        private TokenParser ParseSubsites(ProvisioningHierarchy hierarchy, TokenParser tokenParser, Model.SiteCollection sitecollection, ClientContext siteContext, Web web, TeamNoGroupSubSite subSiteObject)
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

            foreach (var templateRef in sitecollection.Templates)
            {
                var provisioningTemplate = hierarchy.Templates.FirstOrDefault(t => t.Id == templateRef);
                if (provisioningTemplate != null)
                {
                    if (tokenParser == null)
                    {
                        tokenParser = new TokenParser(subweb, provisioningTemplate);
                    }
                    else
                    {
                        tokenParser.Rebase(subweb, provisioningTemplate);
                    }
                    new SiteToTemplateConversion().ApplyRemoteTemplate(subweb, provisioningTemplate, new ProvisioningTemplateApplyingInformation(), true, tokenParser);
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
                    tokenParser = ParseSubsites(hierarchy, tokenParser, sitecollection, siteContext, subweb, subsubSiteObject);
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
