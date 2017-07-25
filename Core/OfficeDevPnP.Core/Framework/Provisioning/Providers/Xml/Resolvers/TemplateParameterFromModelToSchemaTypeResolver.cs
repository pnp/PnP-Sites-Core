using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Template Parameter type from Domain Model to Schema
    /// </summary>
    internal class TemplateParameterFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        private Type _parametersType;

        public TemplateParameterFromModelToSchemaTypeResolver(Type parametersType)
        {
            this._parametersType = parametersType;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var typedSource = source as Dictionary<String, String>;
            var result = Array.CreateInstance(this._parametersType, typedSource.Count);

            if (null != typedSource)
            {
                var index = 0;
                foreach (var l in typedSource)
                {
                    var parameterItem = Activator.CreateInstance(this._parametersType, true);

                    // Extract key and text
                    var key = (String)l.GetType().GetProperty("Key").GetValue(l);
                    var value = (String)l.GetType().GetProperty("Value").GetValue(l);

                    parameterItem.GetType().GetProperty("Key").SetValue(parameterItem, key);
                    parameterItem.GetType().GetProperty("Text").SetValue(parameterItem, new String[] { value });

                    result.SetValue(parameterItem, index);
                    index++;
                }
            }

            return (result);
        }
    }
}
