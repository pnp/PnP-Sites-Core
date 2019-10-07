using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using OfficeDevPnP.Core.Diagnostics;
using System.Text;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Provider for xml based configurations
    /// </summary>
    public abstract class XMLTemplateProvider : TemplateProviderBase
    {
        #region Constructor
        protected XMLTemplateProvider()
            : base()
        {

        }
        protected XMLTemplateProvider(FileConnectorBase connector)
            : base(connector)
        {
        }
        #endregion

        #region Base class overrides

        public override List<ProvisioningTemplate> GetTemplates()
        {
            var formatter = new XMLPnPSchemaFormatter();
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
                if (file.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase))
                {
                    ProvisioningTemplate provisioningTemplate;
                    try
                    {
                        // Use the GetTemplate method to share the same logic
                        provisioningTemplate = this.GetTemplate(file, formatter);
                    }
                    catch (ApplicationException)
                    {
                        Log.Warning(Constants.LOGGING_SOURCE_FRAMEWORK_PROVISIONING, CoreResources.Provisioning_Providers_XML_InvalidFileFormat, file);
                        continue;
                    }

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
            if (uri == null)
            {
                throw new ArgumentNullException(nameof(uri));
            }

            ProvisioningHierarchy result = null;

            var stream = this.Connector.GetFileStream(uri);

            if (stream != null)
            {
                var formatter = new XMLPnPSchemaFormatter();

                ITemplateFormatter specificFormatter = formatter.GetSpecificFormatterInternal(ref stream);
                specificFormatter.Initialize(this);
                result = ((IProvisioningHierarchyFormatter)specificFormatter).ToProvisioningHierarchy(stream);
            }

            return (result);
        }

        public override ProvisioningTemplate GetTemplate(string uri)
        {
            return (this.GetTemplate(uri, (ITemplateProviderExtension[])null));
        }

        public override ProvisioningTemplate GetTemplate(string uri, ITemplateProviderExtension[] extensions = null)
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
            return (this.GetTemplate(uri, identifier, formatter, null));
        }

        public override ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions = null)
        {
            if (String.IsNullOrEmpty(uri))
            {
                throw new ArgumentException("uri");
            }

            if (formatter == null)
            {
                formatter = new XMLPnPSchemaFormatter();
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

            //Resolve xml includes if any
            stream = ResolveXIncludes(stream);

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

        public override ProvisioningTemplate GetTemplate(Stream stream, ITemplateProviderExtension[] extensions = null)
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
            return (this.GetTemplate(stream, identifier, formatter, null));
        }

        public override ProvisioningTemplate GetTemplate(Stream stream, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions = null)
        {
            if (stream == null)
            {
                throw new ArgumentException(nameof(stream));
            }

            if (formatter == null)
            {
                formatter = new XMLPnPSchemaFormatter();
                formatter.Initialize(this);
            }

            // Handle any pre-processing extension
            stream = PreProcessGetTemplateExtensions(extensions, stream);

            //Resolve xml includes if any
            stream = ResolveXIncludes(stream);

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
            this.SaveAs(hierarchy, this.Uri);
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
            if (hierarchy == null)
            {
                throw new ArgumentNullException(nameof(hierarchy));
            }

            if (uri == null)
            {
                throw new ArgumentNullException(nameof(uri));
            }

            if (formatter == null)
            {
                formatter = XMLPnPSchemaFormatter.LatestFormatter;
            }
            formatter.Initialize(this);

            var stream = ((IProvisioningHierarchyFormatter)formatter).ToFormattedHierarchy(hierarchy);

            this.Connector.SaveFileStream(uri, stream);

            if (this.Connector is ICommitableFileConnector)
            {
                ((ICommitableFileConnector)this.Connector).Commit();
            }
        }

        public override void SaveAs(ProvisioningTemplate template, string uri)
        {
            this.SaveAs(template, uri, (ITemplateProviderExtension[])null);
        }

        public override void SaveAs(ProvisioningTemplate template, string uri, ITemplateProviderExtension[] extensions)
        {
            this.SaveAs(template, uri, null, extensions);
        }

        public override void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter)
        {
            this.SaveAs(template, uri, formatter, null);
        }

        public override void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            if (String.IsNullOrEmpty(uri))
            {
                throw new ArgumentException("uri");
            }

            if (formatter == null)
            {
                formatter = new XMLPnPSchemaFormatter();
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

        #region Helper methods

        private Stream ResolveXIncludes(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            var res = stream;
            XDocument xml = XDocument.Load(stream);

            //find XInclude elements by XName
            XName xiName = XName.Get("{http://www.w3.org/2001/XInclude}include");
            var includes = xml.Descendants(xiName).ToList();

            if (includes.Count > 0)
            {
                foreach (var xi in includes)
                {
                    Boolean includeResolved = false;

                    // Resolve xInclude and replace
                    String href = (String)xi.Attribute("href") ?? String.Empty;

                    // If there is the href attribute
                    if (!String.IsNullOrEmpty(href))
                    {
                        Stream incStream = this.Connector.GetFileStream(href);
                        // And if the referenced file can be loaded/resolved
                        if (incStream == null)
                        {
                            //check if include has fallback
                            XName xiFallback = XName.Get("{http://www.w3.org/2001/XInclude}fallback");
                            var fallback = xi.Elements(xiFallback).FirstOrDefault();
                            if ((fallback != null) &&
                                ((fallback.Elements().Count() > 0) || !string.IsNullOrEmpty(fallback.Value)))
                            {
                                var innerXml = fallback.ToXmlElement().InnerXml;
                                incStream = new MemoryStream(Encoding.UTF8.GetBytes(innerXml));
                            }
                        }

                        if (null != incStream)
                        {
                            //resolve include recursive
                            incStream = ResolveXIncludes(incStream);
                            // Replace the xi:include element with the target XML element
                            var resolved = XElement.Load(incStream);
                            xi.ReplaceWith(resolved);
                            includeResolved = true;
                        }
                    }

                    if (!includeResolved)
                    {
                        // Remove the xi:include element 
                        // to avoid any processing failure
                        xi.Remove();
                    }
                }

                //save xml to a new stream
                res = new MemoryStream();
                xml.Save(res);
            }
            res.Seek(0, SeekOrigin.Begin);
            return res;
        }

        #endregion
    }
}
