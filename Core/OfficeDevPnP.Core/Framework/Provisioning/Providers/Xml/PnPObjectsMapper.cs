using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Utility class that maps one object to another
    /// </summary>
    internal static class PnPObjectsMapper
    {
        // TODO: Remember to cover the *Specified problem

        #region MapProperties

        /// <summary>
        /// Maps the properties of a typed source object, to the properties of an untyped destination object
        /// </summary>
        /// <typeparam name="TSource">The type of the source object</typeparam>
        /// <param name="source">The source object</param>
        /// <param name="destination">The destination object</param>
        /// <param name="resolverExpressions">Any custom resolver, optional</param>
        /// <param name="recursive">Defines whether to apply the mapping recursively, optional and by default false</param>
        public static void MapProperties<TSource>(TSource source, Object destination, Dictionary<Expression<Func<TSource, Object>>, IResolver> resolverExpressions = null, Boolean recursive = false)
        {
            Dictionary<string, IResolver> resolvers = ConvertExpressionsToResolvers(resolverExpressions);
            MapProperties(source, destination, resolvers, recursive);
        }

        /// <summary>
        /// Maps the properties of an untyped source object object, to the properties of a typed destination object
        /// </summary>
        /// <typeparam name="TDestination">The type of the destination object</typeparam>
        /// <param name="source">The source object</param>
        /// <param name="destination">The destination object</param>
        /// <param name="resolverExpressions">Any custom resolver, optional</param>
        /// <param name="recursive">Defines whether to apply the mapping recursively, optional and by default false</param>
        public static void MapProperties<TDestination>(Object source, TDestination destination, Dictionary<Expression<Func<TDestination, Object>>, IResolver> resolverExpressions = null, Boolean recursive = false)
        {
            Dictionary<string, IResolver> resolvers = ConvertExpressionsToResolvers(resolverExpressions);
            MapProperties(source, destination, resolvers, recursive);
        }

        /// <summary>
        /// Maps the properties of a source object, to the properties of a destination object
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="destination">The destination object</param>
        /// <param name="resolvers">Any custom resolver, optional</param>
        /// <param name="recursive">Defines whether to apply the mapping recursively, optional and by default false</param>
        public static void MapProperties(Object source, Object destination, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            // Retrieve the list of destination properties
            var destinationProperties = destination?.GetType().GetProperties(
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public);

            // Retrieve the list of source properties
            var sourceProperties = source?.GetType().GetProperties(
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public);

            // Normalize the keys of the resolvers, if any, just in case (maybe this step can be removed)
            if (null != resolvers)
            {
                resolvers = resolvers.ToDictionary(i => i.Key.ToUpper(), i => i.Value);
            }

            // Just for the properties that are not collection or complex types of the model
            // and that are not array or Xml domain model related
            var filteredProperties = destinationProperties?.Where(
                p => (!Attribute.IsDefined(p, typeof(ObsoleteAttribute)) &&
                (p.PropertyType.BaseType.Name != typeof(BaseProvisioningTemplateObjectCollection<>).Name || recursive) &&
                (p.PropertyType.BaseType.Name != typeof(BaseProvisioningHierarchyObjectCollection<>).Name || recursive) &&
                // p.PropertyType.BaseType.Name != typeof(BaseModel).Name && // TODO: Think about this rule ...
                (!p.PropertyType.IsArray || recursive) // &&
                // !p.PropertyType.Namespace.Contains(typeof(XMLConstants).Namespace)
                ));
            foreach (var dp in filteredProperties) // TODO: Think about this rule ...
            {
                // Let's try to see if we have a custom resolver for the current property
                var resolverKey = $"{dp.DeclaringType.FullName}.{dp.Name}".ToUpper();
                var resolver = resolvers != null && resolvers.ContainsKey(resolverKey) ? resolvers[resolverKey] : null;

                // Search for the matching source property
                var sp = sourceProperties?.FirstOrDefault(p => p.Name.Equals(dp.Name, StringComparison.InvariantCultureIgnoreCase));
                var spSpecified = sourceProperties?.FirstOrDefault(p => p.Name.Equals($"{dp.Name}Specified", StringComparison.InvariantCultureIgnoreCase));
                var dpSpecified = destinationProperties?.FirstOrDefault(p => p.Name.Equals($"{dp.Name}Specified", StringComparison.InvariantCultureIgnoreCase));
                if (null != sp || null != resolver)
                {
                    if (null != resolver)
                    {
                        if (resolver is IValueResolver)
                        {
                            // We have a resolver, thus we use it to resolve the input value
                            dp.SetValue(destination, ((IValueResolver)resolver)
                                .Resolve(source, destination, sp?.GetValue(source)));
                        }
                        else if (resolver is ITypeResolver)
                        {
                            // We have a resolver, thus we use it to resolve the input value
                            if (!((ITypeResolver)resolver).CustomCollectionResolver &&
                                (dp.PropertyType.BaseType.Name == typeof(BaseProvisioningTemplateObjectCollection<>).Name ||
                                dp.PropertyType.BaseType.Name == typeof(BaseProvisioningHierarchyObjectCollection<>).Name))
                            {
                                var destinationCollection = dp.GetValue(destination);
                                if (destinationCollection != null)
                                {
                                    var resolvedCollection = ((ITypeResolver)resolver)
                                        .Resolve(source, resolvers, recursive);

                                    destinationCollection.GetType().GetMethod("AddRange",
                                        System.Reflection.BindingFlags.Instance |
                                        System.Reflection.BindingFlags.Public |
                                        System.Reflection.BindingFlags.IgnoreCase)
                                        .Invoke(destinationCollection, new Object[] { resolvedCollection });
                                }
                            }
                            else
                            {
                                dp.SetValue(destination, ((ITypeResolver)resolver)
                                    .Resolve(source, resolvers, recursive));
                            }
                        }
                    }
                    else if (null != sp)
                    {
                        try
                        {
                            // If the destination property is a custom collection of the 
                            // Domain Model and we have the recursive flag enabled
                            if (recursive && (dp.PropertyType.BaseType.Name == typeof(BaseProvisioningTemplateObjectCollection<>).Name ||
                                dp.PropertyType.BaseType.Name == typeof(BaseProvisioningHierarchyObjectCollection<>).Name))
                            {
                                // We need to recursively handle a collection of properties in the Domain Model
                                var destinationCollection = dp.GetValue(destination);
                                if (destinationCollection != null)
                                {
                                    var resolvedCollection =
                                        PnPObjectsMapper.MapObjects(sp.GetValue(source),
                                        new CollectionFromSchemaToModelTypeResolver(
                                            dp.PropertyType.BaseType.GenericTypeArguments[0]), resolvers, recursive);

                                    destinationCollection.GetType().GetMethod("AddRange",
                                        System.Reflection.BindingFlags.Instance |
                                        System.Reflection.BindingFlags.Public |
                                        System.Reflection.BindingFlags.IgnoreCase)
                                        .Invoke(destinationCollection, new Object[] { resolvedCollection });
                                }
                            }
                            // If the destination property is an array of the XML
                            // Schema Model and we have the recursive flag enabled
                            else if (recursive && dp.PropertyType.IsArray)
                            {
                                dp.SetValue(destination,
                                        PnPObjectsMapper.MapObjects(sp.GetValue(source),
                                            new CollectionFromModelToSchemaTypeResolver(dp.PropertyType.IsArray ? dp.PropertyType.GetElementType() : null), 
                                            resolvers, recursive));
                            }
                            else
                            {
                                object sourceValue = sp.GetValue(source);
                                if (sourceValue != null && dp.PropertyType == typeof(string) && sp.PropertyType != typeof(string))
                                {
                                    // Default conversion to String
                                    sourceValue = sourceValue.ToString();
                                }
                                else if (sourceValue != null && dp.PropertyType == typeof(int) && sp.PropertyType != typeof(int))
                                {
                                    // Default conversion to Int32
                                    sourceValue = Int32.Parse(sourceValue.ToString());
                                }
                                else if (sourceValue != null && dp.PropertyType == typeof(bool) && sp.PropertyType != typeof(bool))
                                {
                                    // Default conversion to Boolean
                                    sourceValue = Boolean.Parse(sourceValue.ToString());
                                }
                                else if (sourceValue != null && dp.PropertyType.IsEnum)
                                {
                                    // Default conversion for a target enum type
                                    sourceValue = Enum.Parse(dp.PropertyType, sourceValue.ToString());
                                }
                                else if (sourceValue != null && dp.PropertyType.Name == "Nullable`1" && dp.PropertyType.GenericTypeArguments[0].IsEnum)
                                {
                                    // Default conversion for a target nullable enum type
                                    sourceValue = Enum.Parse(dp.PropertyType.GenericTypeArguments[0], sourceValue.ToString());
                                }
                                else if (sourceValue == null && 
                                    dp.ReflectedType.Namespace == typeof(ProvisioningTemplate).Namespace && 
                                    dp.GetValue(destination) != null)
                                {
                                    // If the destination property is an in memory Domain Model property
                                    // and it has a value, while the source property is null, we keep the
                                    // existing value
                                    sourceValue = dp.GetValue(destination);
                                }
                                else if (sourceValue != null && spSpecified != null)
                                {
                                    // We are processing a property of the schema, which can be nullable
                                    bool isSpecified = (bool)spSpecified.GetValue(source);
                                    if (!isSpecified)
                                    {
                                        sourceValue = null;
                                    }
                                }
                                // We simply need to do 1:1 value mapping
                                dp.SetValue(destination, sourceValue);

                                if (dpSpecified != null)
                                {
                                    dpSpecified.SetValue(destination, true);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Right now, for testing purposes, I just output and skip any issue
                            // TODO: Handle issues insteaf of skipping them, we need to find a common pattern
                        }
                    }
                }
            }
        }

        #endregion

        #region MapObjects

        /// <summary>
        /// Maps a source object, into a destination object
        /// </summary>
        /// <typeparam name="TDestination">The type of the destination object</typeparam>
        /// <param name="source">The source object</param>
        /// <param name="resolver">A custom resolver</param>
        /// <param name="resolverExpressions">Any custom resolver, optional</param>
        /// <param name="recursive">Defines whether to apply the mapping recursively, optional and by default false</param>
        /// <returns>The mapped destination object</returns>
        public static Object MapObjects<TDestination>(Object source, ITypeResolver resolver, Dictionary<Expression<Func<TDestination, Object>>, IResolver> resolverExpressions = null, Boolean recursive = false)
        {
            Dictionary<string, IResolver> resolvers = ConvertExpressionsToResolvers(resolverExpressions);
            return(MapObjects(source, resolver, resolvers, recursive));
        }
        
        /// <summary>
        /// Maps a source object, into a destination object
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="resolver">A custom resolver</param>
        /// <param name="resolvers">Any custom resolver, optional</param>
        /// <param name="recursive">Defines whether to apply the mapping recursively, optional and by default false</param>
        /// <returns>The mapped destination object</returns>
        public static Object MapObjects(Object source, ITypeResolver resolver, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            Object result = null;

            // Normalize the keys of the resolvers, if any
            if (null != resolver)
            {
                result = resolver.Resolve(source, resolvers, recursive);
            }

            return (result);
        }

        #endregion

        #region Utility methods

        /// <summary>
        /// Transforms a Dictionary of IValueResolver instances by Expression into a Dictionary by String (property name)
        /// </summary>
        /// <typeparam name="TTarget">The target Type of the expression</typeparam>
        /// <param name="resolverExpressions">The Dictionary to transform</param>
        /// <returns>The transformed dictionary</returns>
        private static Dictionary<String, IResolver> ConvertExpressionsToResolvers<TTarget>(Dictionary<Expression<Func<TTarget, object>>, IResolver> resolverExpressions)
        {
            Dictionary<String, IResolver> resolvers = null;

            if (resolverExpressions != null)
            {
                resolvers = new Dictionary<String, IResolver>();

                foreach (var re in resolverExpressions.Keys)
                {
                    var propertySelector = re.Body as MemberExpression ?? ((UnaryExpression)re.Body).Operand as MemberExpression;
                    resolvers.Add($"{propertySelector.Member.DeclaringType.FullName}.{propertySelector.Member.Name}".ToUpper(), resolverExpressions[re]);
                }
            }

            return resolvers;
        }

        #endregion
    }
}
