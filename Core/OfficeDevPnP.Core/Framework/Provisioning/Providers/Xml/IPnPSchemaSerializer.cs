using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Basic interface for every Schema Serializer type
    /// </summary>
    public interface IPnPSchemaSerializer
    {
        /// <summary>
        /// Provides the name of the serializer type
        /// </summary>
        String Name { get;  }

        /// <summary>
        /// The method to deserialize an XML Schema based object into a Domain Model object
        /// </summary>
        /// <param name="persistence">The persistence layer object</param>
        /// <param name="template">The PnP Provisioning Template object</param>
        void Deserialize(Object persistence, ProvisioningTemplate template);

        /// <summary>
        /// The method to serialize a Domain Model object into an XML Schema based object 
        /// </summary>
        /// <param name="template">The PnP Provisioning Template object</param>
        /// <param name="persistence">The persistence layer object</param>
        void Serialize(ProvisioningTemplate template, Object persistence);
    }
}
