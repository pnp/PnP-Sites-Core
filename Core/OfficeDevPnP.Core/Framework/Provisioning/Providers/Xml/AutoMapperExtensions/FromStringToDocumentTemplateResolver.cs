using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class FromStringToDocumentTemplateResolver<TContentType, TDocumentTemplate> : IValueResolver<Model.ContentType, TContentType, TDocumentTemplate>
    {
        public TDocumentTemplate Resolve(ContentType source, TContentType destination, TDocumentTemplate destMember, ResolutionContext context)
        {
            // TODO: Implement this resolver
            return (default(TDocumentTemplate));
        }
    }
}
