using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Json
{
    public class JsonPnPFormatter : ITemplateFormatterWithValidation
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        public bool IsValid(Stream template)
        {
            return GetValidationResults(template).IsValid;
        }

        public ValidationResult GetValidationResults(System.IO.Stream template)
        {
            // We do not provide JSON validation capabilities
            return new ValidationResult { IsValid = true, Exceptions = null };
        }

        public System.IO.Stream ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            String jsonString = JsonConvert.SerializeObject(template, new BasePermissionsConverter());
            Byte[] jsonBytes = System.Text.Encoding.Unicode.GetBytes(jsonString);
            MemoryStream jsonStream = new MemoryStream(jsonBytes);
            jsonStream.Position = 0;

            return (jsonStream);
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template, string identifier)
        {
            StreamReader sr = new StreamReader(template, Encoding.Unicode);
            String jsonString = sr.ReadToEnd();
            Model.ProvisioningTemplate result = JsonConvert.DeserializeObject<Model.ProvisioningTemplate>(jsonString, new BasePermissionsConverter());
            return (result);
        }
    }
}
