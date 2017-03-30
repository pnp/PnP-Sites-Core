using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    internal class ExpressionValueResolver<T> : IValueResolver
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        private Func<T, object> expression = null;

        public ExpressionValueResolver(Func<T, object> expression)
        {
            this.expression = expression;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            return expression.Invoke((T)sourceValue);
        }
    }

    internal class ExpressionValueResolver : IValueResolver 
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        private Func<object, object, object> expression = null;

        public ExpressionValueResolver(Func<object, object, object> expression)
        {
            if(expression == null)
            {
                throw new ArgumentException("expression");
            }
            this.expression = expression;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            return expression.Invoke(source, sourceValue);
        }
    }
}
