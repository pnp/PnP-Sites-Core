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
        public string Name => this.GetType().Name;

        private Func<T, object, object, object> expression1 = null;
        private Func<object, T, object, object> expression2 = null;
        private Func<object, object, T, object> expression3 = null;


        public ExpressionValueResolver(Func<object, object, T, object> expression)
        {
            this.expression3 = expression;
        }

        public ExpressionValueResolver(Func<T, object> expression)
        {
            this.expression2 = (source, sourceValue, destination) => expression.Invoke(sourceValue);
        }

        public ExpressionValueResolver(Func<T, object, object> expression)
        {
            this.expression1 = (source, sourceValue, destination) => expression.Invoke(source, sourceValue);
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            object res = null;
            if (expression1 != null)
            {
                res = expression1.Invoke((T)source, sourceValue, destination);
            }
            else if (expression2 != null)
            {
                res = expression2.Invoke(source, (T)sourceValue, destination);
            }
            else if (expression3 != null)
            {
                res = expression3.Invoke(source, sourceValue, (T)destination);
            }
            return res;
        }
    }

    internal class ExpressionValueResolver : IValueResolver 
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        private Func<object, object, object, object> expression = null;

        public ExpressionValueResolver(Func<object, object, object, object> expression)
        {
            if (expression == null)
            {
                throw new ArgumentException("expression");
            }
            this.expression = expression;
        }

        public ExpressionValueResolver(Func<object, object, object> expression)
        {
            if(expression == null)
            {
                throw new ArgumentException("expression");
            }
            this.expression = (source, sourceValue, destination) => expression.Invoke(source, sourceValue);
        }

        public ExpressionValueResolver(Func<object> expression)
        {
            if (expression == null)
            {
                throw new ArgumentException("expression");
            }
            this.expression = (source, sourceValue, destination) => expression.Invoke();
        }


        public object Resolve(object source, object destination, object sourceValue)
        {
            return expression.Invoke(source, sourceValue, destination);
        }
    }
}
