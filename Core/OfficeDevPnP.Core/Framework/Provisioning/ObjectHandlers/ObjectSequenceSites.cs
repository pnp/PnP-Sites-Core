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
    internal class ObjectSequenceSites : ObjectSequenceHandlerBase
    {
        public override string Name => "Sequences";

        public override ProvisioningTemplate ExtractObjects(Tenant tenant, Model.Provisioning template, ProvisioningTemplateCreationInformation creationInfo)
        {
            throw new NotImplementedException();
        }

        public override TokenParser ProvisionObjects(Tenant tenant, Model.Provisioning sequenceTemplate, TokenParser tokenParser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {
                foreach (var sequence in sequenceTemplate.Sequences)
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
                                        Alias = t.Alias,
                                        DisplayName = t.Title,
                                        Description = t.Description,
                                        Classification = t.Classification,
                                        IsPublic = t.IsPublic
                                    };
                                    var groupSiteInfo = Sites.SiteCollection.GetGroupInfo(tenant.Context as ClientContext, t.Alias).GetAwaiter().GetResult();
                                    if (groupSiteInfo == null)
                                    {
                                        WriteMessage($"Creating Team Site {t.Alias}", ProvisioningMessageType.Progress);
                                        siteContext = Sites.SiteCollection.CreateAsync(tenant.Context as ClientContext, siteInfo).GetAwaiter().GetResult();
                                    }
                                    else
                                    {
                                        if (groupSiteInfo.ContainsKey("siteUrl"))
                                        {
                                            WriteMessage($"Using existing Team Site {t.Alias}", ProvisioningMessageType.Progress);
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
                                        Classification = c.Classification,
                                        Description = c.Description,
                                        Lcid = (uint)c.Language,
                                        Owner = c.Owner,
                                        Title = c.Title,
                                        Url = c.Url
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
                                    if (tenant.SiteExists(c.Url))
                                    {
                                        WriteMessage($"Using existing Communications Site {c.Url}", ProvisioningMessageType.Progress);
                                        siteContext = (tenant.Context as ClientContext).Clone(c.Url);
                                    }
                                    else
                                    {
                                        WriteMessage($"Creating Communications Site {c.Url}", ProvisioningMessageType.Progress);
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
                                        Title = t.Title,
                                        Url = t.Url,
                                        SiteOwnerLogin = t.Owner,
                                    };
                                    WriteMessage($"Creating Team Site with no Office 365 group {t.Url}", ProvisioningMessageType.Progress);
                                    if (tenant.SiteExists(t.Url))
                                    {
                                        WriteMessage($"Using existing Team Site {t.Url}", ProvisioningMessageType.Progress);
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
                            var provisioningTemplate = sequenceTemplate.Templates.FirstOrDefault(t => t.Id == templateRef);
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
                                new SiteToTemplateConversion().ApplyRemoteTemplate(web, provisioningTemplate, new ProvisioningTemplateApplyingInformation(), tokenParser);
                            }
                            else
                            {
                                WriteMessage($"Referenced template ID {templateRef} not found", ProvisioningMessageType.Error);
                            }

                        }

                        foreach (var subsite in sitecollection.Sites)
                        {
                            var subSiteObject = (TeamNoGroupSubSite)subsite;
                            web.EnsureProperties(w => w.Webs.IncludeWithDefaultProperties(), w => w.ServerRelativeUrl);

                            //var managedPath = web.ServerRelativeUrl.ToLower().Contains("/sites/") ? "sites" : "teams";

                            var subweb = web.Webs.FirstOrDefault(t => t.ServerRelativeUrl.Equals(UrlUtility.Combine(web.ServerRelativeUrl, "/", subSiteObject.Url.Trim(new char[] { '/' }))));
                            if (subweb == null)
                            {
                                subweb = web.Webs.Add(new WebCreationInformation()
                                {
                                    Language = int.Parse(subSiteObject.Language),
                                    Url = subSiteObject.Url,
                                    Description = subSiteObject.Description,
                                    Title = subSiteObject.Title,
                                    UseSamePermissionsAsParentSite = subSiteObject.UseSamePermissionsAsParentSite,
                                    WebTemplate = "STS#3"
                                });
                                siteContext.Load(subweb);
                                siteContext.ExecuteQueryRetry();
                            }

                            foreach (var templateRef in sitecollection.Templates)
                            {
                                var provisioningTemplate = sequenceTemplate.Templates.FirstOrDefault(t => t.Id == templateRef);
                                if (provisioningTemplate != null)
                                {
                                    if (siteTokenParser == null)
                                    {
                                        siteTokenParser = new TokenParser(subweb, provisioningTemplate);
                                    }
                                    else
                                    {
                                        siteTokenParser.Rebase(subweb, provisioningTemplate);
                                    }
                                    new SiteToTemplateConversion().ApplyRemoteTemplate(subweb, provisioningTemplate, new ProvisioningTemplateApplyingInformation(), tokenParser);
                                }
                                else
                                {
                                    WriteMessage($"Referenced template ID {templateRef} not found", ProvisioningMessageType.Error);
                                }
                            }
                        }
                    }
                }
                return tokenParser;
            }
        }

        public override bool WillExtract(Tenant tenant, Model.Provisioning sequenceTemplate, ProvisioningTemplateCreationInformation creationInfo)
        {
            throw new NotImplementedException();
        }

        public override bool WillProvision(Tenant tenant, Model.Provisioning sequenceTemplate, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return sequenceTemplate.Sequences.Count > 0;
        }
    }
}
