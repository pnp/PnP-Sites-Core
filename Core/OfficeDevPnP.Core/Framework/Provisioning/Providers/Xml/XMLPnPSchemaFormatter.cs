using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Helper class that abstracts from any specific version of XMLPnPSchemaFormatter
    /// </summary>
    public class XMLPnPSchemaFormatter : ITemplateFormatter
    {
        private TemplateProviderBase _provider;

        public void Initialize(TemplateProviderBase provider)
        {
            this._provider = provider;
        }

        #region Static methods and properties

        /// <summary>
        /// Static property to retrieve an instance of the latest XMLPnPSchemaFormatter
        /// </summary>
        public static ITemplateFormatter LatestFormatter
        {
            get
            {
                return (new XMLPnPSchemaV201805Serializer());
            }
        }

        /// <summary>
        /// Static method to retrieve a specific XMLPnPSchemaFormatter instance
        /// </summary>
        /// <param name="version"></param>
        /// <returns></returns>
        public static ITemplateFormatter GetSpecificFormatter(XMLPnPSchemaVersion version)
        {
            switch (version)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                case XMLPnPSchemaVersion.V201503:
                    return (new XMLPnPSchemaV201503Formatter());
                case XMLPnPSchemaVersion.V201505:
                    return (new XMLPnPSchemaV201505Formatter());
                case XMLPnPSchemaVersion.V201508:
                    return (new XMLPnPSchemaV201508Formatter());
#pragma warning restore CS0618 // Type or member is obsolete
                case XMLPnPSchemaVersion.V201512:
                    return (new XMLPnPSchemaV201512Formatter());
                case XMLPnPSchemaVersion.V201605:
                    return (new XMLPnPSchemaV201605Formatter());
                case XMLPnPSchemaVersion.V201705:
                    return (new XMLPnPSchemaV201705Serializer());
                case XMLPnPSchemaVersion.V201801:
                    return (new XMLPnPSchemaV201801Serializer());
                case XMLPnPSchemaVersion.V201805:
                default:
                    return (new XMLPnPSchemaV201805Serializer());
            }
        }

        /// <summary>
        /// Static method to retrieve a specific XMLPnPSchemaFormatter instance
        /// </summary>
        /// <param name="namespaceUri"></param>
        /// <returns></returns>
        public static ITemplateFormatter GetSpecificFormatter(string namespaceUri)
        {
            switch (namespaceUri)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_03:
                    return (new XMLPnPSchemaV201503Formatter());
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_05:
                    return (new XMLPnPSchemaV201505Formatter());
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_08:
                    return (new XMLPnPSchemaV201508Formatter());
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12:
                    return (new XMLPnPSchemaV201512Formatter());
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05:
                    return (new XMLPnPSchemaV201605Formatter());
#pragma warning restore CS0618 // Type or member is obsolete
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05:
                    return (new XMLPnPSchemaV201705Serializer());
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2018_01:
                    return (new XMLPnPSchemaV201801Serializer());
                case XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2018_05:
                default:
                    return (new XMLPnPSchemaV201805Serializer());
            }
        }

        #endregion

        #region Abstract methods implementation

        public bool IsValid(System.IO.Stream template)
        {
            ITemplateFormatter formatter = this.GetSpecificFormatterInternal(ref template);
            formatter.Initialize(this._provider);
            return (formatter.IsValid(template));
        }

        public System.IO.Stream ToFormattedTemplate(Model.ProvisioningTemplate template)
        {
            ITemplateFormatter formatter = XMLPnPSchemaFormatter.LatestFormatter;
            formatter.Initialize(this._provider);
            return (formatter.ToFormattedTemplate(template));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template)
        {
            return (this.ToProvisioningTemplate(template, null));
        }

        public Model.ProvisioningTemplate ToProvisioningTemplate(System.IO.Stream template, String identifier)
        {
            ITemplateFormatter formatter = this.GetSpecificFormatterInternal(ref template);
            formatter.Initialize(this._provider);
            return (formatter.ToProvisioningTemplate(template, identifier));
        }

        #endregion

        #region Helper Methods

        private ITemplateFormatter GetSpecificFormatterInternal(ref System.IO.Stream template)
        {
            if (template == null)
            {
                throw new ArgumentNullException(nameof(template));
            }

            // Create a copy of the source stream
            MemoryStream sourceStream = new MemoryStream();
            template.Position = 0;
            template.CopyTo(sourceStream);
            sourceStream.Position = 0;
            template = sourceStream;

            XDocument xml = XDocument.Load(template);
            template.Position = 0;

            String targetNamespaceUri = xml.Root.Name.NamespaceName;

            if (!String.IsNullOrEmpty(targetNamespaceUri))
            {
                return (XMLPnPSchemaFormatter.GetSpecificFormatter(targetNamespaceUri));
            }
            else
            {
                return (XMLPnPSchemaFormatter.LatestFormatter);
            }
        }

        #endregion
    }
}

