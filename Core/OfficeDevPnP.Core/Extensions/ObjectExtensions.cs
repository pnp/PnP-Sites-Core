using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Extensions
{
    /// <summary>
    /// Provide general purpose extension methods
    /// </summary>
    public static class ObjectExtensions
    {

        /// <summary>
        /// Set an object field or property and returns if the value was changed.
        /// </summary>
        /// <typeparam name="TObject">Type of the target object</typeparam>
        /// <typeparam name="T">T of the property</typeparam>
        /// <param name="target">target object </param>
        /// <param name="propertyToSet">Expression to the property or field of the object</param>
        /// <param name="valueToSet">new value to set</param>
        /// <param name="allowNull">continue with set operation is null value is specified</param>
        /// <param name="allowEmpty">continue with set operation is null or empty value is specified</param>
        /// <returns><c>true</c> if the value has changed, otherwise <c>false</c></returns>
        public static bool Set<TObject, T>(this TObject target, Expression<Func<TObject, T>> propertyToSet, T valueToSet, bool allowNull=true, bool allowEmpty=true)
        {
            // Taken from https://stackoverflow.com/a/29092675/588868
            var members = new List<MemberInfo>();

            var exp = propertyToSet.Body;

            if (!allowNull && valueToSet == null)
            {
                return false;
            }

            if (!allowEmpty && (valueToSet is string) && (valueToSet == null || (valueToSet as string) == ""))
            {
                return false;
            }

            while (exp != null)
            {
                var mi = exp as MemberExpression;

                if (mi != null)
                {
                    members.Add(mi.Member);
                    exp = mi.Expression;
                }
                else
                {
                    var pe = exp as ParameterExpression;

                    if (pe == null)
                    {
                        // We support only a ParameterExpression at the base
                        throw new NotSupportedException();
                    }

                    break;
                }
            }

            if (members.Count == 0)
            {
                // We need at least a getter
                throw new NotSupportedException();
            }

            // Now we must walk the getters (excluding the last).
            object targetObject = target;

            // We have to walk the getters from last (most inner) to second
            // (the first one is the one we have to use as a setter)
            for (int i = members.Count - 1; i >= 1; i--)
            {
                var pi = members[i] as PropertyInfo;

                if (pi != null)
                {
                    targetObject = pi.GetValue(targetObject);
                }
                else
                {
                    var fi = (FieldInfo)members[i];
                    targetObject = fi.GetValue(targetObject);
                }
            }

            // The first one is the getter we treat as a setter
            {
                var pi = members[0] as PropertyInfo;

                if (pi != null)
                {
                    var current = (T)pi.GetValue(targetObject);
                    if (!EqualityComparer<T>.Default.Equals(current, valueToSet))
                    {
                        pi.SetValue(targetObject, valueToSet);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    var fi = (FieldInfo)members[0];
                    var current = (T)fi.GetValue(targetObject);
                    if (!EqualityComparer<T>.Default.Equals(current, valueToSet))
                    {
                        fi.SetValue(targetObject, valueToSet);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }

        /// <summary>
        /// Nullify a string when it's an empty one
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string NullIfEmpty(this string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }

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
