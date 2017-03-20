using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Decimal value into a Double
    /// </summary>
    internal class ExpressionTypeResolver<T> : ITypeResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        private Action<T, object> expression = null;

        private Type targetItemType;

        public ExpressionTypeResolver(Type targetItemType, Action<T, object> expression)
        {
            this.targetItemType = targetItemType;
            this.expression = expression;
        }

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            if (null != source)
            {
                var result = Activator.CreateInstance(this.targetItemType, true);
                expression.Invoke((T)source, result);
                return (result);
            }
            else
            {
                return (null);
            }
        }
    }
}
