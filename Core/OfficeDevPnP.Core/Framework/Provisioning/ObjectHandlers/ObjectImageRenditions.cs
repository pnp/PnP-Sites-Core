using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Feature = OfficeDevPnP.Core.Framework.Provisioning.Model.Feature;
using System;
using System.Linq;
using OfficeDevPnP.Core.Diagnostics;
using Microsoft.SharePoint.Client.Publishing;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectImageRenditions : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Image Renditions"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                var site = (web.Context as ClientContext).Site;

                // Check if this is not a noscript site as publishing features are not supported
                if (web.IsNoScriptSite())
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Publishing_SkipProvisioning);
                    return parser;
                }

                var webFeatureActive = web.IsFeatureActive(Constants.FeatureId_Web_Publishing);
                var siteFeatureActive = site.IsFeatureActive(Constants.FeatureId_Site_Publishing);
                if (template.Publishing.AutoCheckRequirements == AutoCheckRequirementsOptions.SkipIfNotCompliant && !webFeatureActive)
                {
                    scope.LogDebug("Publishing Feature (Web Scoped) not active. Skipping provisioning of Publishing Image Renditions");
                    return parser;
                }
                else if (template.Publishing.AutoCheckRequirements == AutoCheckRequirementsOptions.MakeCompliant)
                {
                    if (!siteFeatureActive)
                    {
                        scope.LogDebug("Making site compliant for publishing");
                        site.ActivateFeature(Constants.FeatureId_Site_Publishing);
                        web.ActivateFeature(Constants.FeatureId_Web_Publishing);
                    }
                    else
                    {
                        if (!web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                        {
                            scope.LogDebug("Making site compliant for publishing");
                            web.ActivateFeature(Constants.FeatureId_Web_Publishing);
                        }
                    }
                }
                else if (!webFeatureActive)
                {
                    throw new Exception("Publishing Feature not active. Provisioning failed");
                }

                if (template.Publishing != null && 
                    template.Publishing.ImageRenditions != null && 
                    template.Publishing.ImageRenditions.Count > 0)
                {
                    var renditions = SiteImageRenditions.GetRenditions(context);
                    context.ExecuteQueryRetry();

                    foreach (var r in template.Publishing.ImageRenditions)
                    {
                        var rendition = new Microsoft.SharePoint.Client.Publishing.ImageRendition();
                        rendition.Name = r.Name;
                        rendition.Height = r.Height;
                        rendition.Width = r.Width;

                        if (!renditions.Any(rd => rd.Name == rendition.Name && rd.Height == rendition.Height && rd.Width == rendition.Width))
                        {
                            renditions.Add(rendition);
                        }
                    }

                    SiteImageRenditions.SetRenditions(context, renditions);
                    context.ExecuteQueryRetry();
                }
            }
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;

                // We extract the Image Renditions if and only if the Publishing feature is enabled
                if (web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                {
                    // This Object Handler will be invoked after the Publishing handler
                    // Thus, we should have the Publishing property assigned in the template
                    if (template.Publishing == null)
                    {
                        // And if not, we create it
                        template.Publishing = new Publishing();
                    }

                    var renditions = SiteImageRenditions.GetRenditions(context);
                    context.ExecuteQueryRetry();

                    foreach (var r in renditions)
                    {
                        template.Publishing.ImageRenditions.Add(new Model.ImageRendition
                        {
                            Name = r.Name,
                            Height = r.Height,
                            Width = r.Width,
                        });
                    }
                }
            }
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Publishing != null && template.Publishing.ImageRenditions.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }
}
