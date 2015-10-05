using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Xml.Linq;
using OfficeDevPnP.Core.Entities;
using System.IO;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPublishing : ObjectHandlerBase
    {
        private const string AVAILABLEPAGELAYOUTS = "__PageLayouts";
        private const string DEFAULTPAGELAYOUT = "__DefaultPageLayout";
        private readonly Guid PUBLISHING_FEATURE_WEB = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
        private readonly Guid PUBLISHING_FEATURE_SITE = new Guid("f6924d36-2fa8-4f0b-b16d-06b7250180fa");

        public override string Name
        {
            get { return "Publishing"; }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (web.IsFeatureActive(PUBLISHING_FEATURE_WEB))
                {
                    web.EnsureProperty(w => w.Language);
                    var webTemplates = web.GetAvailableWebTemplates(web.Language, false);
                    web.Context.Load(webTemplates, wts => wts.Include(wt => wt.Name, wt => wt.Lcid));
                    web.Context.ExecuteQueryRetry();
                    Publishing publishing = new Publishing();
                    publishing.AvailableWebTemplates.AddRange(webTemplates.AsEnumerable<WebTemplate>().Select(wt => new AvailableWebTemplate() { TemplateName = wt.Name, LanguageCode = (int)wt.Lcid }));
                    publishing.AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant;
                    publishing.DesignPackage = null;
                    publishing.PageLayouts.AddRange(GetAvailablePageLayouts(web));
                    template.Publishing = publishing;
                }
            }
            return template;
        }


        private IEnumerable<PageLayout> GetAvailablePageLayouts(Web web)
        {
            var defaultLayoutXml = web.GetPropertyBagValueString(DEFAULTPAGELAYOUT, null);

            var defaultPageLayoutUrl = string.Empty;
            if (defaultLayoutXml != null && defaultLayoutXml != "__inherit")
            {
                defaultPageLayoutUrl = XElement.Parse(defaultLayoutXml).Attribute("url").Value;
            }

            List<PageLayout> layouts = new List<PageLayout>();

            var layoutsXml = web.GetPropertyBagValueString(AVAILABLEPAGELAYOUTS, null);

            if (!string.IsNullOrEmpty(layoutsXml) && layoutsXml != "__inherit")
            {
                var layoutsElement = XElement.Parse(layoutsXml);

                foreach (var layout in layoutsElement.Descendants("layout"))
                {
                    if (layout.Attribute("url") != null)
                    {
                        var pageLayout = new PageLayout();
                        pageLayout.Path = layout.Attribute("url").Value;

                        if (pageLayout.Path == defaultPageLayoutUrl)
                        {
                            pageLayout.IsDefault = true;
                        }
                        layouts.Add(pageLayout);
                    }

                }
            }
            return layouts;
        }




        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var site = (web.Context as ClientContext).Site;

                var webFeatureActive = web.IsFeatureActive(PUBLISHING_FEATURE_WEB);
                var siteFeatureActive = site.IsFeatureActive(PUBLISHING_FEATURE_SITE);
                if (template.Publishing.AutoCheckRequirements == AutoCheckRequirementsOptions.SkipIfNotCompliant && !webFeatureActive)
                {
                    scope.LogDebug("Publishing Feature (Web Scoped) not active. Skipping provisioning of Publishing settings");
                    return parser;
                }
                else if (template.Publishing.AutoCheckRequirements == AutoCheckRequirementsOptions.MakeCompliant)
                {
                    if (!siteFeatureActive)
                    {
                        scope.LogDebug("Making site compliant for publishing");
                        site.ActivateFeature(PUBLISHING_FEATURE_SITE);
                        web.ActivateFeature(PUBLISHING_FEATURE_WEB);
                    }
                    else
                    {
                        if (!web.IsFeatureActive(PUBLISHING_FEATURE_WEB))
                        {
                            scope.LogDebug("Making site compliant for publishing");
                            web.ActivateFeature(PUBLISHING_FEATURE_WEB);
                        }
                    }
                }
                else
                {
                    throw new Exception("Publishing Feature not active. Provisioning failed");
                }

                var availableWebTemplates = template.Publishing.AvailableWebTemplates.Select(t => new WebTemplateEntity() { LanguageCode = t.LanguageCode.ToString(), TemplateName = t.TemplateName }).ToList();
                if (availableWebTemplates.Any())
                {
                    web.SetAvailableWebTemplates(availableWebTemplates);
                }
                var availablePageLayouts = template.Publishing.PageLayouts.Select(p => p.Path);
                if (availablePageLayouts.Any())
                {
                    web.SetAvailablePageLayouts(site.RootWeb, availablePageLayouts);
                }
                if (template.Publishing.DesignPackage != null)
                {
                    var package = template.Publishing.DesignPackage;

                    var tempFileName = Path.Combine(Path.GetTempPath(), template.Connector.GetFilenamePart(package.DesignPackagePath));
                    scope.LogDebug("Saving {0} to temporary file: {1}", package.DesignPackagePath, tempFileName);
                    using (var stream = template.Connector.GetFileStream(package.DesignPackagePath))
                    {
                        using (var outstream = System.IO.File.Create(tempFileName))
                        {
                            stream.CopyTo(outstream);
                        }
                    }
                    scope.LogDebug("Installing design package");
                    site.InstallSolution(package.PackageGuid, tempFileName, package.MajorVersion, package.MinorVersion);
                    System.IO.File.Delete(tempFileName);
                }
                return parser;
            }
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return web.IsFeatureActive(PUBLISHING_FEATURE_WEB);
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            return template.Publishing != null;
        }
    }
}
