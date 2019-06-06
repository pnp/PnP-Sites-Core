using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Base class for every Schema Serializer
    /// </summary>
    internal abstract class PnPBaseSchemaSerializer<TModelType> : IPnPSchemaSerializer
    {
        public string Name
        {
            get { return (this.GetType().Name); }
        }

        public abstract void Deserialize(object persistence, ProvisioningTemplate template);

        public abstract void Serialize(ProvisioningTemplate template, object persistence);

        /// <summary>
        /// Protected method to create a Lambda Expression like: i => i.Property
        /// </summary>
        /// <param name="targetType">The Type of the .NET property to apply the Lambda Expression to</param>
        /// <param name="propertyName">The name of the property of the target object</param>
        /// <returns></returns>
        protected LambdaExpression CreateSelectorLambda(Type targetType, String propertyName)
        {
            return (Expression.Lambda(
                Expression.Convert(
                    Expression.MakeMemberAccess(
                        Expression.Parameter(targetType, "i"),
                        targetType.GetProperty(propertyName,
                            System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.Public)),
                    typeof(object)),
                ParameterExpression.Parameter(targetType, "i")));
        }
    }
}
