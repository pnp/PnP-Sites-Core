using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201705
{
    internal class ListInstanceDataRowsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public ListInstanceDataRowsFromModelToSchemaTypeResolver()
        {
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            var list = source as Model.DataRowCollection;

            if (null != list)
            {
                //var listInstanceViewsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ListInstanceViews, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                //var listInstanceViewsType = Type.GetType(listInstanceViewsTypeName, true);
                //result = Activator.CreateInstance(listInstanceViewsType);

                //result.GetPublicInstanceProperty("RemoveExistingViews").SetValue(result, list.RemoveExistingViews);

                //var xmlElements = new List<XmlElement>();
                //foreach (var view in list.Views)
                //{
                //    var viewXml = XElement.Parse(view.SchemaXml);
                //    xmlElements.Add(viewXml.ToXmlElement());
                //}

                //var anyElements = result.GetPublicInstanceProperty("Any");
                //anyElements.SetValue(result, xmlElements.ToArray());
            }

            return (result);
        }
    }
}
