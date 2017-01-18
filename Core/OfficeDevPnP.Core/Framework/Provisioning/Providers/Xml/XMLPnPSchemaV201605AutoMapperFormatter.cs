using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Schema;
using AutoMapper;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperProfiles;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    internal class XMLPnPSchemaV201605AutoMapperFormatter :
        XMLPnPSchemaAutoMapperFormatter
    {
        public XMLPnPSchemaV201605AutoMapperFormatter(): base(
            XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05, 
            typeof(XMLPnPSchemaAutoMapperFormatter)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2016-05.xsd"))
        {

        }

        protected override IMapper CreateMapperForFormattedTemplate()
        {
            // AutoMapper configuration
            var config = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile(new V201605Profile());
            });

            return (config.CreateMapper());
        }

        protected override IMapper CreateMapperForProvisioningTemplate()
        {
            // AutoMapper configuration
            var config = new MapperConfiguration(cfg =>
            {
                cfg.AddProfile(new V201605Profile());
            });

            return (config.CreateMapper());
        }
    }
}
