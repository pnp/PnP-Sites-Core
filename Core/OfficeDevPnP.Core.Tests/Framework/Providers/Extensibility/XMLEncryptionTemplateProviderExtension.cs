using OfficeDevPnP.Core.Framework.Provisioning.Providers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;

#if !NETSTANDARD2_0
namespace OfficeDevPnP.Core.Tests.Framework.Providers.Extensibility
{
    public class XMLEncryptionTemplateProviderExtension : ITemplateProviderExtension
    {
        public bool SupportsGetTemplatePostProcessing
        {
            get
            {
                return (false);
            }
        }

        public bool SupportsGetTemplatePreProcessing
        {
            get
            {
                return (true);
            }
        }

        public bool SupportsSaveTemplatePostProcessing
        {
            get
            {
                return (true);
            }
        }

        public bool SupportsSaveTemplatePreProcessing
        {
            get
            {
                return (false);
            }
        }

        private X509Certificate2 _certificate;

        public void Initialize(object settings)
        {
            _certificate = settings as X509Certificate2;
        }

        public ProvisioningTemplate PostProcessGetTemplate(ProvisioningTemplate template)
        {
            throw new NotImplementedException();
        }

        public Stream PostProcessSaveTemplate(Stream stream)
        {
            MemoryStream result = new MemoryStream();

            var namespaces = new Dictionary<string, string>();
            namespaces.Add("pnp", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05);

            SecureXml.EncryptXmlDocument(stream, result, "/pnp:Provisioning", namespaces, this._certificate);
            result.Position = 0;

            return (result);
        }

        public Stream PreProcessGetTemplate(Stream stream)
        {
            MemoryStream result = new MemoryStream();

            SecureXml.DecryptXmlDocument(stream, result, this._certificate);
            result.Position = 0;

            return (result);
        }

        public ProvisioningTemplate PreProcessSaveTemplate(ProvisioningTemplate template)
        {
            throw new NotImplementedException();
        }
    }
}
#endif