using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    [SupportedTemplateSchemas(Schemas = SupportedSchema.V201605 | SupportedSchema.V201703)]
    public class WebSettingsSerializer : ISchemaSerializer
    {
        public XElement FromProvisioningTemplate(ProvisioningTemplate template, XNamespace ns)
        {
            var element = new XElement(ns + "WebSettings");
            if (template.WebSettings == null) return element;
            element.AddOptionalAttribute(nameof(WebSettings.Title), template.WebSettings.Title);
            element.AddOptionalAttribute(nameof(WebSettings.RequestAccessEmail), template.WebSettings.RequestAccessEmail);
            element.AddOptionalAttribute(nameof(WebSettings.NoCrawl), template.WebSettings.NoCrawl);
            element.AddOptionalAttribute(nameof(WebSettings.WelcomePage), template.WebSettings.WelcomePage);
            element.AddOptionalAttribute(nameof(WebSettings.Description), template.WebSettings.Description);
            element.AddOptionalAttribute(nameof(WebSettings.SiteLogo), template.WebSettings.SiteLogo);
            element.AddOptionalAttribute(nameof(WebSettings.AlternateCSS), template.WebSettings.AlternateCSS);
            element.AddOptionalAttribute(nameof(WebSettings.MasterPageUrl), template.WebSettings.MasterPageUrl);
            element.AddOptionalAttribute(nameof(WebSettings.CustomMasterPageUrl), template.WebSettings.CustomMasterPageUrl);
            return element;
        }

        public ProvisioningTemplate ToProvisioningTemplate(XElement templateElement, XNamespace ns, ProvisioningTemplate template)
        {
            var webSettings = templateElement.Descendants(ns + "WebSettings").FirstOrDefault();
            if (webSettings == null) return template;

            template.WebSettings = new WebSettings
            {
                Title = webSettings.Attribute(nameof(WebSettings.Title))?.Value,
                RequestAccessEmail = webSettings.Attribute(nameof(WebSettings.RequestAccessEmail))?.Value,
                NoCrawl = webSettings.GetOptionalBoolValue(nameof(WebSettings.NoCrawl)),
                WelcomePage = webSettings.Attribute(nameof(WebSettings.WelcomePage))?.Value,
                Description = webSettings.Attribute(nameof(WebSettings.Description))?.Value,
                SiteLogo = webSettings.Attribute(nameof(WebSettings.SiteLogo))?.Value,
                AlternateCSS = webSettings.Attribute(nameof(WebSettings.AlternateCSS))?.Value,
                MasterPageUrl = webSettings.Attribute(nameof(WebSettings.MasterPageUrl))?.Value,
                CustomMasterPageUrl = webSettings.Attribute(nameof(template.WebSettings.CustomMasterPageUrl))?.Value
            };
            return template;
        }
    }
}
