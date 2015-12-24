using System;
using System.IO;
using System.Reflection;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using System.Linq;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// This class will be used to provide access to the right base template configuration
    /// </summary>
    public static class BaseTemplateManager
    {
        private static readonly Guid PUBLISHING_FEATURE_WEB = new Guid("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");

        public static ProvisioningTemplate GetBaseTemplate(this Web web)
        {
            web.Context.Load(web, p => p.WebTemplate, p => p.Configuration);
            web.Context.ExecuteQueryRetry();

            //if (web.IsFeatureActive(PUBLISHING_FEATURE_WEB) && web.WebTemplate == "STS" && web.Configuration == 0)
            //{
            //    return GetBaseTemplate(web, "STS0PUBLISHING", 0);
            //}
            //else
            //{
            return GetBaseTemplate(web, web.WebTemplate, web.Configuration);
            //}
        }

        public static ProvisioningTemplate GetBaseTemplate(this Web web, string webTemplate, short configuration)
        {

            ProvisioningTemplate provisioningTemplate = null;

            try
            {
                string baseTemplate = string.Format("OfficeDevPnP.Core.Framework.Provisioning.BaseTemplates.v{0}.{1}{2}Template.xml", GetSharePointVersion(), webTemplate, configuration);
                using (Stream stream = typeof(BaseTemplateManager).Assembly.GetManifestResourceStream(baseTemplate))
                {
                    // Figure out the formatter to use
                    XDocument z = XDocument.Load(stream);
                    var result = z.Root.Attributes().Where(a => a.IsNamespaceDeclaration).
                            GroupBy(a => a.Name.Namespace == XNamespace.None ? String.Empty : a.Name.LocalName,
                                    a => XNamespace.Get(a.Value)).
                            ToDictionary(g => g.Key,
                                         g => g.First());
                    var pnpns = result["pnp"];

                    stream.Seek(0, SeekOrigin.Begin);
                    // Get the XML document from the stream
                    ITemplateFormatter formatter = XMLPnPSchemaFormatter.GetSpecificFormatter(pnpns.NamespaceName);

                    // And convert it into a ProvisioningTemplate

                    provisioningTemplate = formatter.ToProvisioningTemplate(stream);
                }
            }
            catch (Exception ex)
            {
                OfficeDevPnP.Core.Diagnostics.Log.Error(ex, "Provisioning", "Error occured while retrieving basetemplate");
            }

            return provisioningTemplate;
        }


        private static string GetSharePointVersion()
        {
            Assembly asm = Assembly.GetAssembly(typeof(Site));
            AssemblyName name = asm.GetName();
            return String.Format("{0}_{1}", name.Version.Major, name.Version.Minor);
        }

    }
}
