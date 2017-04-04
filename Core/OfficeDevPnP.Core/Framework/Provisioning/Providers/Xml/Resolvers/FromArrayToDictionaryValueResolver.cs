using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves an Array of object into a Dictionary
    /// </summary>
    internal class FromArrayToDictionaryValueResolver<TKey, TValue> : IValueResolver
    {
        public string Name => this.GetType().Name;

        private String _keyField;
        private String _valueField;
        private Type _sourceArrayItemType;
        private String _sourcePropertyName;

        public FromArrayToDictionaryValueResolver(Type sourceArrayItemType,
            LambdaExpression keySelector, LambdaExpression valueSelector, String sourcePropertyName = null)
        {
            this._sourceArrayItemType = sourceArrayItemType;

            var keyField = keySelector.Body as MemberExpression ?? ((UnaryExpression)keySelector.Body).Operand as MemberExpression;
            var valueField = valueSelector.Body as MemberExpression ?? ((UnaryExpression)valueSelector.Body).Operand as MemberExpression;

            this._keyField = keyField.Member.Name;
            this._valueField = valueField.Member.Name;
            this._sourcePropertyName = sourcePropertyName;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            var result = new Dictionary<TKey, TValue>();

            if (null == sourceValue && null != source && !string.IsNullOrEmpty(_sourcePropertyName))
            {
                //get source value from property with non-matching name
                sourceValue = source.GetPublicInstancePropertyValue(_sourcePropertyName);
            }
            else if (null == sourceValue && null != source)
            {
                // If we don't have the source value, but we have the source object
                // we try to retrieve the source value from the source object
                sourceValue = source.GetType().GetProperties().FirstOrDefault(sp => sp.PropertyType.IsArray &&
                    sp.PropertyType.GetElementType().FullName == _sourceArrayItemType.FullName)?.GetValue(source);
            }

            if (null != sourceValue)
            {
                foreach (var l in (IEnumerable)sourceValue)
                {
                    // Extract key and text
                    var key = (TKey)l.GetType().GetProperty(this._keyField).GetValue(l);
                    var value = (TValue)l.GetType().GetProperty(this._valueField).GetValue(l);
                    result.Add(key, value);
                }
            }

            return (result);
        }
    }
}
