using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for Security from Schema to Model
    /// </summary>
    internal class TemplateSecurityUsersFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        private String collectionName;

        public TemplateSecurityUsersFromSchemaToModelTypeResolver(String collectionName)
        {
            this.collectionName = collectionName;
        }

        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            var result = new List<Model.User>();

            var userCollection = source.GetPublicInstancePropertyValue(this.collectionName);
            if (null != userCollection)
            {
                var userResolver = new CollectionFromSchemaToModelTypeResolver(typeof(Model.User));
                result.AddRange(userResolver.Resolve(userCollection.GetPublicInstancePropertyValue("User"))
                    as IEnumerable<Model.User>);
            }

            return (result);
        }
    }
}
