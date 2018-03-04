using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{

    internal class ExpressionTypeResolver<T> : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;

        private Action<T, Dictionary<String, IResolver>, bool, object> expression = null;

        private Type targetItemType;

        public ExpressionTypeResolver(Type targetItemType, Action<T, object> expression)
        {
            this.targetItemType = targetItemType;
            this.expression = (source, resolvers, recursive, result) => expression.Invoke(source, result);
        }

        public ExpressionTypeResolver(Type targetItemType, Action<T, Dictionary<String, IResolver>, bool, object> expression)
        {
            this.targetItemType = targetItemType;
            this.expression = expression;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            if (null != source)
            {
                var result = Activator.CreateInstance(this.targetItemType, true);
                expression.Invoke((T)source, resolvers, recursive, result);
                return (result);
            }
            else
            {
                return (null);
            }
        }
    }
}
