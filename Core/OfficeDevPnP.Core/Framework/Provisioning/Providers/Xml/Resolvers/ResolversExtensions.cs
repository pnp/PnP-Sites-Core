using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Extension type for resolvers
    /// </summary>
    public static class ResolversExtensions
    {
        /// <summary>
        /// Retrieves the value of a public, instance property 
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="propertyName">The property name, case insensitive</param>
        /// <returns>The property value, if any</returns>
        public static Object GetPublicInstancePropertyValue(this Object source, String propertyName)
        {
            return (source?.GetType()?.GetProperty(propertyName,
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase)?
                .GetValue(source));
        }

        /// <summary>
        /// Retrieves a public, instance property 
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="propertyName">The property name, case insensitive</param>
        /// <returns>The property, if any</returns>
        public static PropertyInfo GetPublicInstanceProperty(this Object source, String propertyName)
        {
            return (source?.GetType()?.GetProperty(propertyName,
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase));
        }

        /// <summary>
        /// Sets the value of a public, instance property 
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="propertyName">The property name, case insensitive</param>
        /// <param name="value">The value to set</param>
        public static void SetPublicInstancePropertyValue(this Object source, String propertyName, object value)
        {
            source?.GetType()?.GetProperty(propertyName,
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public |
                System.Reflection.BindingFlags.IgnoreCase)?
                .SetValue(source, value);
        }
    }
}
