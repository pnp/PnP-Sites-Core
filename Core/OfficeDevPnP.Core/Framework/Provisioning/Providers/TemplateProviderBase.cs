using System;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System.IO;
using System.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers
{
    public abstract class TemplateProviderBase
    {
        private Dictionary<string, string> _parameters = new Dictionary<string, string>();
        private bool _supportSave = false;
        private bool _supportDelete = false;
        private FileConnectorBase _connector = null;
        private string _uri = "";

        #region Constructors
        
        public TemplateProviderBase()
        {

        }

        public TemplateProviderBase(FileConnectorBase connector)
        {
            this._connector = connector;
        }

        #endregion

        #region Public Properties

        public Dictionary<string, string> Parameters
        {
            get
            {
                return this._parameters;
            }
        }

        public virtual bool SupportsSave
        {
            get
            {
                return _supportSave;
            }
        }

        public virtual bool SupportsDelete
        {
            get
            {
                return _supportDelete;
            }
        }

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

        public abstract List<ProvisioningTemplate> GetTemplates();

        public abstract List<ProvisioningTemplate> GetTemplates(ITemplateFormatter formatter);

        public abstract ProvisioningTemplate GetTemplate(string uri);

        public abstract ProvisioningTemplate GetTemplate(string uri, ITemplateProviderExtension[] extensions);

        public abstract ProvisioningTemplate GetTemplate(string uri, string identifier);

        public abstract ProvisioningTemplate GetTemplate(string uri, ITemplateFormatter formatter);

        public abstract ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter);

        public abstract ProvisioningTemplate GetTemplate(string uri, string identifier, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);
        
        public abstract void Save(ProvisioningTemplate template);

        public abstract void Save(ProvisioningTemplate template, ITemplateProviderExtension[] extensions);

        public abstract void Save(ProvisioningTemplate template, ITemplateFormatter formatter);

        public abstract void Save(ProvisioningTemplate template, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);

        public abstract void SaveAs(ProvisioningTemplate template, string uri);

        public abstract void SaveAs(ProvisioningTemplate template, string uri, ITemplateProviderExtension[] extensions);

        public abstract void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter);

        public abstract void SaveAs(ProvisioningTemplate template, string uri, ITemplateFormatter formatter, ITemplateProviderExtension[] extensions);

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
