using AutoMapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.AutoMapperExtensions
{
    /// <summary>
    /// Defines some custom mapping behaviors common to all the type maps
    /// </summary>
    public static class MappingExtensions
    {
        private const String ProvisioningTemplateSchemaNamespaceTrailer = "OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml";
        private const String SpecifiedSuffix = "Specified";

        /// <summary>
        /// Defines the default behavior for every map
        /// </summary>
        /// <param name="map">The source map</param>
        /// <returns>The updated map</returns>
        public static IMappingExpression ApplyDefaultMappingBehavior(this IMappingExpression map)
        {
            return(map.ReverseMap().AfterMap((source, destination) => 
            {
                // Get the types underlying the mapping objects
                var sourceType = source.GetType();
                var destinationType = destination.GetType();

                // Get all the instance public properties of the source and destination objects
                var sourceProperties = sourceType.GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
                var destinationProperties = destinationType.GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);

                // Let's see if we are mapping from the Domain Model to the XML Schema model
                if (destinationType.Namespace.StartsWith(ProvisioningTemplateSchemaNamespaceTrailer))
                {
                    foreach (var property in destinationProperties)
                    {
                        // If we have a *Specified property in the XML Schema model type we need to handle it
                        if (property.Name.EndsWith(SpecifiedSuffix, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // Look for a corresponding property without the *Specified suffix
                            var referenceDestinationProperty = destinationProperties.FirstOrDefault(p => p.Name == property.Name.Substring(0, property.Name.Length - SpecifiedSuffix.Length));

                            // If we have such a property, let's dig into it's mapping property to determine
                            // the value to assign to the *Specified property in the XML Schema model
                            if (referenceDestinationProperty != null)
                            {
                                // If the corresponding property in the source object exists
                                var sourceProperty = sourceProperties.FirstOrDefault(p => p.Name == referenceDestinationProperty.Name);
                                if (sourceProperty != null)
                                {
                                    // And if it is of type Nullable<T>
                                    if (System.Nullable.GetUnderlyingType(sourceProperty.PropertyType) != null)
                                    {
                                        // We need to evaluate the HasValue property in order to define the *Specified property
                                        var hasValueOutcome = false;
                                        var hasValue = sourceProperty.PropertyType.GetProperty("HasValue", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
                                        var nullablePropertyValue = sourceProperty.GetValue(source);
                                        if (nullablePropertyValue != null)
                                        {
                                            hasValueOutcome = (Boolean)hasValue.GetValue(nullablePropertyValue);
                                        }

                                        // Set the *Specified value to the outcome
                                        property.SetValue(destination, hasValueOutcome);
                                    }
                                    else
                                    {
                                        // Otherwise we simply need to set the *Specified property if
                                        // the source property value does not equal default(T), where 
                                        // T is the type of the source property 
                                        property.SetValue(destination, sourceProperty.GetValue(source) != null);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // We are mapping from the XML Schema model to the Domain Model
                }
            }));
        }
    }
}
