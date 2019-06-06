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
    /// Resolve collection from model to schema with expression
    /// </summary>
    internal class ExpressionCollectionValueResolver : IValueResolver
    {
        public string Name => this.GetType().Name;

        private Type targetItemType = null;
        private Func<object, object> expression = null;

        public ExpressionCollectionValueResolver(Expression<Func<object, object>> expression, Type targetItemType)
        {
            if (targetItemType == null)
            {
                throw new ArgumentException("targetItemType");
            }
            if (expression == null)
            {
                throw new ArgumentException("expression");
            }
            this.expression = expression.Compile();
            this.targetItemType = targetItemType;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            object result = null;

            if ((null != sourceValue)&&(sourceValue is IList))
            {
                var sourceList = (IList)sourceValue;
                var resultArray = Array.CreateInstance(this.targetItemType, sourceList.Count);

                int index = 0;
                foreach (var i in (IEnumerable)sourceValue)
                {
                    var targetItem = this.expression.Invoke(i);
                    resultArray.SetValue(targetItem, index++);
                }
                result = resultArray;
            }
            return (result);
        }
    }

    /// <summary>
    /// Resolve collection from schema to model with expression
    /// </summary>
    internal class ExpressionCollectionValueResolver<T> : IValueResolver 
    {
        public string Name => this.GetType().Name;

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
