using System;
using System.Collections;
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
    internal class ExpressionCollectionValueResolver<T> : IValueResolver 
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        private Func<object, T> expression = null;

        public ExpressionCollectionValueResolver(Expression<Func<object, T>> expression)
        {
            if(expression == null)
            {
                throw new ArgumentException("expression");
            }
            this.expression = expression.Compile();
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            var result = new List<T>();

            if (null != sourceValue)
            {
                foreach (var i in (IEnumerable)sourceValue)
                {
                    var targetItem = this.expression.Invoke(i);
                    result.Add(targetItem);
                }
            }
            return (result);
        }
    }
}
