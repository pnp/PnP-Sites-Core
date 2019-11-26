using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Json
{
    /// <summary>
    /// Provider for JSON based configurations
    /// </summary>
    public abstract class JsonTemplateProvider : TemplateProviderBase
    {
        #region Constructor
        protected JsonTemplateProvider() : base()
        {

        }

        protected JsonTemplateProvider(FileConnectorBase connector)
            : base(connector)
        {
        }
        #endregion

        #region Base class overrides

        public override List<ProvisioningTemplate> GetTemplates()
        {
            var formatter = new JsonPnPFormatter();
            formatter.Initialize(this);
            return (this.GetTemplates(formatter));
        }

        public override List<ProvisioningTemplate> GetTemplates(ITemplateFormatter formatter)
        {
            List<ProvisioningTemplate> result = new List<ProvisioningTemplate>();

            // Retrieve the list of available template files
            List<String> files = this.Connector.GetFiles();

            // For each file
            foreach (var file in files)
            {
                if (file.EndsWith(".json", StringComparison.InvariantCultureIgnoreCase))
                {
                    // And convert it into a ProvisioningTemplate
                    ProvisioningTemplate provisioningTemplate = this.GetTemplate(file, formatter);

                    if (provisioningTemplate != null)
                    {
                        // Add the template to the result
                        result.Add(provisioningTemplate);
                    }
                }
            }

            return (result);
        }

        public override ProvisioningHierarchy GetHierarchy(string uri)
        {
            throw new NotImplementedException();
        }

        public override ProvisioningTemplate GetTemplate(string uri)
        {
            return (this.GetTemplate(uri, (ITemplateProviderExtension[])null));
        }

        public override ProvisioningTemplate GetTemplate(string uri, ITemplateProviderExtension[] extensions)
        {
            return (this.GetTemplate(uri, null, null, extensions));
        }

        public override ProvisioningTemplate GetTemplate(string uri, string identifier)
        {
            return (this.GetTemplate(uri, identifier, null));
        }

        public override ProvisioningTemplate GetTemplate(string uri, ITemplateFormatter formatter)
        {
            return (this.GetTemplate(uri, null, formatter));
        }

        public override ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter)
        {
            return (this.GetTemplate(uri, null, formatter, null));
        }

        public override ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions)
        {
            if (String.IsNullOrEmpty(uri))
            {
                throw new ArgumentException("uri");
            }

            if (formatter == null)
            {
                formatter = new JsonPnPFormatter();
                formatter.Initialize(this);
            }

            // Get the XML document from a File Stream
            Stream stream = this.Connector.GetFileStream(uri);

            if (stream == null)
            {
                throw new ApplicationException(string.Format(CoreResources.Provisioning_Formatter_Invalid_Template_URI, uri));
            }

            // Handle any pre-processing extension
            stream = PreProcessGetTemplateExtensions(extensions, stream);

            // And convert it into a ProvisioningTemplate
            ProvisioningTemplate provisioningTemplate = formatter.ToProvisioningTemplate(stream, identifier);

            // Handle any post-processing extension
            provisioningTemplate = PostProcessGetTemplateExtensions(extensions, provisioningTemplate);

            // Store the identifier of this template, is needed for latter save operation
            this.Uri = uri;

            return (provisioningTemplate);
        }

        public override ProvisioningTemplate GetTemplate(Stream stream)
        {
            return (this.GetTemplate(stream, (ITemplateProviderExtension[])null));
        }

        public override ProvisioningTemplate GetTemplate(Stream stream, ITemplateProviderExtension[] extensions)
        {
            return (this.GetTemplate(stream, null, null, extensions));
        }

        public override ProvisioningTemplate GetTemplate(Stream stream, string identifier)
        {
            return (this.GetTemplate(stream, identifier, null));
        }

        public override ProvisioningTemplate GetTemplate(Stream stream, ITemplateFormatter formatter)
        {
            return (this.GetTemplate(stream, null, formatter));
        }

        public override ProvisioningTemplate GetTemplate(Stream stream, string identifier, ITemplateFormatter formatter)
        {
            return (this.GetTemplate(stream, null, formatter, null));
        }

        public override ProvisioningTemplate GetTemplate(Stream stream, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions)
        {
            if (stream == null)
            {
                throw new ArgumentException(nameof(stream));
            }

            if (formatter == null)
            {
                formatter = new JsonPnPFormatter();
                formatter.Initialize(this);
            }

            // Handle any pre-processing extension
            stream = PreProcessGetTemplateExtensions(extensions, stream);

            // And convert it into a ProvisioningTemplate
            ProvisioningTemplate provisioningTemplate = formatter.ToProvisioningTemplate(stream, identifier);

            // Handle any post-processing extension
            provisioningTemplate = PostProcessGetTemplateExtensions(extensions, provisioningTemplate);

            // Store the identifier of this template, is needed for latter save operation
            this.Uri = null;

            return (provisioningTemplate);
        }

        public override void Save(ProvisioningHierarchy hierarchy)
        {
            throw new NotImplementedException();
        }

        public override void Save(ProvisioningTemplate template)
        {
            this.Save(template, (ITemplateProviderExtension[])null);
        }

        public override void Save(ProvisioningTemplate template, ITemplateProviderExtension[] extensions = null)
        {
            this.Save(template, null, extensions);
        }

        public override void Save(ProvisioningTemplate template, ITemplateFormatter formatter)
        {
            this.Save(template, formatter, null);
        }

        public override void Save(ProvisioningTemplate template, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions = null)
        {
            this.SaveAs(template, this.Uri, formatter, extensions);
        }

        public override void SaveAs(ProvisioningHierarchy hierarchy, string uri, ITemplateFormatter formatter = null)
        {
            throw new NotImplementedException();
        }

        public override void SaveAs(ProvisioningTemplate template, string uri)
        {
            this.SaveAs(template, uri, (ITemplateProviderExtension[])null);
        }

        public override void SaveAs(ProvisioningTemplate template, string uri, ITemplateProviderExtension[] extensions = null)
        {
            this.SaveAs(template, uri, null, extensions);
        }

        public override void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter)
        {
            this.SaveAs(template, uri, formatter, null);
        }

        public override void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions = null)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            if (String.IsNullOrEmpty(uri))
            {
                throw new ArgumentException(nameof(uri));
            }

            if (formatter == null)
            {
                formatter = new JsonPnPFormatter();
            }

            SaveToConnector(template, uri, formatter, extensions);
        }

        public override void Delete(string uri)
        {
            if (String.IsNullOrEmpty(uri))
            {
                throw new ArgumentException("identifier");
            }

            this.Connector.DeleteFile(uri);
        }
       
        #endregion
    }
}
