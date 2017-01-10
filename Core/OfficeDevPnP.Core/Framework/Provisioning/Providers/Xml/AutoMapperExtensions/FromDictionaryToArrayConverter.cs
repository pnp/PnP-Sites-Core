using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    public class FromDictionaryToArrayConverter<TKey, TValue, TItem> : ITypeConverter<Dictionary<TKey, TValue>, TItem[]>
        where TItem : class, new()
    {
        private String _keyField;
        private String _valueField;

        public FromDictionaryToArrayConverter(Expression<Func<TItem, object>> keySelector, Expression<Func<TItem, object>> valueSelector)
        {
            var keyField = keySelector.Body as MemberExpression ?? ((UnaryExpression)keySelector.Body).Operand as MemberExpression;
            var valueField = valueSelector.Body as MemberExpression ?? ((UnaryExpression)valueSelector.Body).Operand as MemberExpression;

            this._keyField = keyField.Member.Name;
            this._valueField = valueField.Member.Name;
        }

        public TItem[] Convert(Dictionary<TKey, TValue> source, TItem[] destination, ResolutionContext context)
        {
            var result = new List<TItem>();

            foreach (var item in source)
            {
                var resultItem = new TItem();
                resultItem.GetType().GetProperty(this._keyField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public).SetValue(resultItem, item.Key);
                resultItem.GetType().GetProperty(this._valueField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public).SetValue(resultItem, item.Value);
                result.Add(resultItem);
            }

            return (result.ToArray<TItem>());
        }
    }
}
