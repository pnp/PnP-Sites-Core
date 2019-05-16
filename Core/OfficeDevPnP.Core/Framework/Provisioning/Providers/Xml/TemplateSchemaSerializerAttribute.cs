using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Attribute for Template Schema Serializers
    /// </summary>
    public class TemplateSchemaSerializerAttribute : Attribute
    {
        /// <summary>
        /// The schemas supported by the serializer
        /// </summary>
        public XMLPnPSchemaVersion MinimalSupportedSchemaVersion { get; set; }

        /// <summary>
        /// The sequence number for applying the serializer during serialization
        /// </summary>
        /// <remarks>
        /// Should be a multiple of 100, to make room for future new insertions
        /// </remarks>
        public Int32 SerializationSequence { get; set; } = 0;

        /// <summary>
        /// The sequence number for applying the serializer during deserialization
        /// </summary>
        /// <remarks>
        /// Should be a multiple of 100, to make room for future new insertions
        /// </remarks>
        public Int32 DeserializationSequence { get; set; } = 0;

        /// <summary>
        /// Defines the scope of the serializer
        /// </summary>
        /// <remarks>
        /// By default the serializers target a single Provisioning Template
        /// </remarks>
        public SerializerScope Scope { get; set; } = SerializerScope.ProvisioningTemplate;
    }

    /// <summary>
    /// Defines the scope of a serializer
    /// </summary>
    public enum SerializerScope
    {
        /// <summary>
        /// The serializer targets a single Provisioning Template
        /// </summary>
        ProvisioningTemplate,
        /// <summary>
        /// The serializer targets a full Provisioning file but not a tenant Template
        /// </summary>
        Provisioning,
        /// <summary>
        /// The serializer targets a Provisioning Hierarchy
        /// </summary>
        ProvisioningHierarchy,
        /// <summary>
        /// The serializer targets the whole Tenant
        /// </summary>
        Tenant,
    }
}
