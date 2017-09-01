using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Collections;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves a Dictionary into an Array of objects
    /// </summary>
    internal class FromDictionaryToArrayValueResolver<TKey, TValue> : IValueResolver
    {
        public string Name => this.GetType().Name;

        private String _keyField;
        private String _valueField;
        private Type _targetArrayItemType;
        private String _sourcePropertyName;

        public FromDictionaryToArrayValueResolver(Type targetArrayItemType,
            LambdaExpression keySelector, LambdaExpression valueSelector, string sourcePropertyName = null)
        {
            this._targetArrayItemType = targetArrayItemType;

            var keyField = keySelector.Body as MemberExpression ?? ((UnaryExpression)keySelector.Body).Operand as MemberExpression;
            var valueField = valueSelector.Body as MemberExpression ?? ((UnaryExpression)valueSelector.Body).Operand as MemberExpression;

            this._keyField = keyField.Member.Name;
            this._valueField = valueField.Member.Name;
            this._sourcePropertyName = sourcePropertyName;
        }

        public object Resolve(object source, object destination, object sourceValue)
        {
            object result = null;

            if (null == sourceValue && null != source && !string.IsNullOrEmpty(_sourcePropertyName))
            {
                //get source value from property with non-matching name
                sourceValue = source.GetPublicInstancePropertyValue(_sourcePropertyName);
            }

            var sourceDictionary = sourceValue != null && sourceValue is IEnumerable<KeyValuePair<TKey, TValue>> ?
                sourceValue as IEnumerable<KeyValuePair<TKey, TValue>>:
                source as IEnumerable<KeyValuePair<TKey, TValue>>;

            if (null == sourceDictionary && null != sourceValue)
            {
                throw new ArgumentException("Invalid source object. Expected type implementing IEnumerable<KeyValuePair<TKey, TValue>>", "source");
            }
            else if (null != sourceDictionary && sourceDictionary.Count() > 0)
            {
                var listType = typeof(List<>);
                var resultType = this._targetArrayItemType.MakeArrayType();

                var resultArray = (Array)Activator.CreateInstance(resultType, sourceDictionary.Count());
                var i = 0;
                foreach (var item in sourceDictionary)
                {
                    var resultItem = Activator.CreateInstance(this._targetArrayItemType);
                    resultItem.GetType().GetProperty(this._keyField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public).SetValue(resultItem, item.Key);
                    resultItem.GetType().GetProperty(this._valueField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public).SetValue(resultItem, item.Value);
                    resultArray.SetValue(resultItem, i++);
                }
                result = resultArray;
            }

            return (result);
        }
    }
}
