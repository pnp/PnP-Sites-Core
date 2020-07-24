using System;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System.IO;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers
{
    /// <summary>
    /// Handles methods for Template Provider
    /// </summary>
    public abstract class TemplateProviderBase
    {
        private Dictionary<string, string> _parameters = new Dictionary<string, string>();
        private bool _supportSave = false;
        private bool _supportDelete = false;
        private FileConnectorBase _connector = null;
        private string _uri = "";

        #region Constructors
        /// <summary>
        /// Default Constructor
        /// </summary>
        public TemplateProviderBase()
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="connector"></param>
        public TemplateProviderBase(FileConnectorBase connector)
        {
            this._connector = connector;
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// Template parameters
        /// </summary>
        public Dictionary<string, string> Parameters
        {
            get
            {
                return this._parameters;
            }
        }

        /// <summary>
        /// Supports template save
        /// </summary>
        public virtual bool SupportsSave
        {
            get
            {
                return _supportSave;
            }
        }

        /// <summary>
        /// Supports template delete
        /// </summary>
        public virtual bool SupportsDelete
        {
            get
            {
                return _supportDelete;
            }
        }

        /// <summary>
        /// File Connector
        /// </summary>
        public virtual FileConnectorBase Connector
        {
            get
            {
                return _connector;
            }
            set
            {
                _connector = value;
            }
        }

        /// <summary>
        /// Uri of site
        /// </summary>
        public String Uri
        {
            get
            {
                return _uri;
            }
            set
            {
                _uri = value;
            }
        }

        #endregion

        #region Abstract Methods
        /// <summary>
        /// Gets list of ProvisioningTemplates
        /// </summary>
        /// <returns>Returns collection of ProvisioningTemplate</returns>
        public abstract List<ProvisioningTemplate> GetTemplates();

        /// <summary>
        /// Gets list of ProvisioningTemplates
        /// </summary>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <returns>Returns collection of ProvisioningTemplate</returns>
        public abstract List<ProvisioningTemplate> GetTemplates(ITemplateFormatter formatter);

        /// <summary>
        /// Gets ProvisioningHierarchy
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <returns>Returns a ProvisioningHierarchy</returns>
        public abstract ProvisioningHierarchy GetHierarchy(string uri);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(string uri);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(string uri, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <param name="identifier">ProvisioningTemplate identifier</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(string uri, string identifier);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(string uri, ITemplateFormatter formatter);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <param name="identifier">ProvisioningTemplate identifier</param>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The source uri</param>
        /// <param name="identifier">ProvisioningTemplate identifier</param>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(Stream stream);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(Stream stream, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <param name="identifier">ProvisioningTemplate identifier</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(Stream stream, string identifier);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(Stream stream, ITemplateFormatter formatter);

        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <param name="identifier">ProvisioningTemplate identifier</param>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(Stream stream, string identifier, ITemplateFormatter formatter);


        /// <summary>
        /// Gets ProvisioningTemplate
        /// </summary>
        /// <param name="stream">The source stream</param>
        /// <param name="identifier">ProvisioningTemplate identifier</param>
        /// <param name="formatter">Provisioning Template formatter</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        /// <returns>Returns a ProvisioningTemplate</returns>
        public abstract ProvisioningTemplate GetTemplate(Stream stream, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Saves ProvisioningHierarchy
        /// </summary>
        /// <param name="hierarchy">Provisioning Hierarchy</param>
        public abstract void Save(ProvisioningHierarchy hierarchy);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        public abstract void Save(ProvisioningTemplate template);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        public abstract void Save(ProvisioningTemplate template, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="formatter">Provisioning Template Formatter</param>
        public abstract void Save(ProvisioningTemplate template, ITemplateFormatter formatter);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="formatter">Provisioning Template Formatter</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        public abstract void Save(ProvisioningTemplate template, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Saves ProvisioningHierarchy
        /// </summary>
        /// <param name="hierarchy">Provisioning Hierarchy</param>
        /// <param name="uri">The target uri</param>
        /// <param name="formatter">Provisioning Template Formatter</param>
        public abstract void SaveAs(ProvisioningHierarchy hierarchy, string uri, ITemplateFormatter formatter = null);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="uri">The target uri</param>
        public abstract void SaveAs(ProvisioningTemplate template, string uri);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="uri">The target uri</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        public abstract void SaveAs(ProvisioningTemplate template, string uri, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="uri">The target uri</param>
        /// <param name="formatter">Provisioning Template Formatter</param>
        public abstract void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter);

        /// <summary>
        /// Saves ProvisioningTemplate
        /// </summary>
        /// <param name="template">Provisioning Template</param>
        /// <param name="uri">The target uri</param>
        /// <param name="formatter">Provisioning Template Formatter</param>
        /// <param name="extensions">Collection of provisioning template extensions</param>
        public abstract void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);

        /// <summary>
        /// Deletes ProvisioningTemplate
        /// </summary>
        /// <param name="uri">The target uri</param>
        public abstract void Delete(string uri);

        #endregion

        #region Protected methods

        protected virtual void SaveToConnector(ProvisioningTemplate template, string uri, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions)
        {
            if (String.IsNullOrEmpty(template.Id))
            {
                template.Id = Path.GetFileNameWithoutExtension(uri);
            }

            template = PreProcessSaveTemplateExtensions(extensions, template);

            using (var stream = formatter.ToFormattedTemplate(template))
            {
                using (var processedStream = PostProcessSaveTemplateExtensions(extensions, stream))
                {
                    this.Connector.SaveFileStream(uri, processedStream);
                }
            }

            if (this.Connector is ICommitableFileConnector)
            {
                ((ICommitableFileConnector)this.Connector).Commit();
            }
        }

        /// <summary>
        /// This method is invoked before calling the formatter to serialize the template
        /// </summary>
        /// <param name="extensions">The list of custom extensions</param>
        /// <param name="template">The template to serialize</param>
        /// <returns>The template eventually updated by the custom extensions</returns>
        protected virtual ProvisioningTemplate PreProcessSaveTemplateExtensions(ITemplateProviderExtension[] extensions, ProvisioningTemplate template)
        {
            ProvisioningTemplate result = template;

            // Handle any pre-processing extension during Save
            if (extensions != null && extensions.Length > 0)
            {
                foreach (var extension in extensions.Where(e => e.SupportsSaveTemplatePreProcessing))
                {
                    result = extension.PreProcessSaveTemplate(result);
                }
            }

            return (result);
        }

        protected Stream PostProcessSaveTemplateExtensions(ITemplateProviderExtension[] extensions, Stream stream)
        {
            // Handle any pre-processing extension
            if (extensions != null && extensions.Length > 0)
            {
                foreach (var extension in extensions.Where(e => e.SupportsSaveTemplatePostProcessing))
                {
                    var temp = new MemoryStream();
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(temp);

                    temp.Seek(0, SeekOrigin.Begin);
                    stream = extension.PostProcessSaveTemplate(temp);
                    stream.Seek(0, SeekOrigin.Begin);
                }
            }

            return (stream);
        }

        protected Stream PreProcessGetTemplateExtensions(ITemplateProviderExtension[] extensions, Stream stream)
        {
            // Handle any pre-processing extension
            if (extensions != null && extensions.Length > 0)
            {
                foreach (var extension in extensions.Where(e => e.SupportsGetTemplatePreProcessing))
                {
                    var temp = new MemoryStream();
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(temp);

                    temp.Seek(0, SeekOrigin.Begin);
                    stream = extension.PreProcessGetTemplate(temp);
                    stream.Seek(0, SeekOrigin.Begin);
                }
            }

            return (stream);
        }

        protected ProvisioningTemplate PostProcessGetTemplateExtensions(ITemplateProviderExtension[] extensions, ProvisioningTemplate template)
        {
            ProvisioningTemplate result = template;

            // Handle any post-processing extension
            if (extensions != null && extensions.Length > 0)
            {
                foreach (var extension in extensions.Where(e => e.SupportsGetTemplatePostProcessing))
                {
                    result = extension.PostProcessGetTemplate(result);
                }
            }

            return (result);
        }

        #endregion
    }
}
