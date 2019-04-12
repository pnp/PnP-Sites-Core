using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201801;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201805;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers.V201807;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers
{
    /// <summary>
    /// Class to serialize/deserialize the Tenant-wide settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201807,
        SerializationSequence = 100, DeserializationSequence = 100,
        Scope = SerializerScope.ProvisioningHierarchy)]
    internal class SequenceSerializer : PnPBaseSchemaSerializer<ProvisioningSequence>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var sequences = persistence.GetPublicInstancePropertyValue("Sequence");

            if (sequences != null)
            {
                var expressions = new Dictionary<Expression<Func<ProvisioningSequence, Object>>, IResolver>();

                // Handle the TermStore property of the Sequence, if any
                expressions.Add(seq => seq.TermStore, new ExpressionValueResolver((s, v) => {

                    if (v != null)
                    {
                        var tgs = new TermGroupsSerializer();
                        var termGroupsExpressions = tgs.GetTermGroupDeserializeExpressions();

                        var result = new Model.ProvisioningTermStore();
                        result.TermGroups.AddRange(
                            PnPObjectsMapper.MapObjects<TermGroup>(v,
                                new CollectionFromSchemaToModelTypeResolver(typeof(TermGroup)),
                                termGroupsExpressions,
                                recursive: true)
                                as IEnumerable<TermGroup>);

                        return (result);
                    }
                    else
                    {
                        return (null);
                    }
                }));

                // Handle the SiteCollections property of the Sequence, if any
                expressions.Add(seq => seq.SiteCollections, 
                    new SiteCollectionsAndSitesFromSchemaToModelTypeResolver(typeof(SiteCollection)));
                expressions.Add(seq => seq.SiteCollections[0].Sites,
                    new SiteCollectionsAndSitesFromSchemaToModelTypeResolver(typeof(SubSite)));
                expressions.Add(seq => seq.SiteCollections[0].Sites[0].Sites,
                    new SiteCollectionsAndSitesFromSchemaToModelTypeResolver(typeof(SubSite)));
                expressions.Add(seq => seq.SiteCollections[0].Templates, new ExpressionValueResolver((s, v) => {

                    var result = new List<String>();

                    if (v != null)
                    {
                        foreach (var t in (IEnumerable)v)
                        {
                            var templateId = t.GetPublicInstancePropertyValue("ID")?.ToString();

                            if (templateId != null)
                            {
                                result.Add(templateId);
                            }
                        }
                    }

                    return (result);
                }));
                expressions.Add(seq => seq.SiteCollections[0].Sites[0].Templates, new ExpressionValueResolver((s, v) => {

                    var result = new List<String>();

                    if (v != null)
                    {
                        foreach (var t in (IEnumerable)v)
                        {
                            var templateId = t.GetPublicInstancePropertyValue("ID")?.ToString();

                            if (templateId != null)
                            {
                                result.Add(templateId);
                            }
                        }
                    }

                    return (result);
                }));

                template.ParentHierarchy.Sequences.AddRange(
                PnPObjectsMapper.MapObjects<ProvisioningSequence>(sequences,
                        new CollectionFromSchemaToModelTypeResolver(typeof(ProvisioningSequence)),
                        expressions, 
                        recursive: true)
                        as IEnumerable<ProvisioningSequence>);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ParentHierarchy != null && 
                template.ParentHierarchy.Sequences != null &&
                template.ParentHierarchy.Sequences.Count > 0)
            {
                var sequenceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Sequence, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var sequenceType = Type.GetType(sequenceTypeName, true);

                var expressions = new Dictionary<string, IResolver>();

                // Handle the TermStore property of the Sequence, if any
                expressions.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Sequence.TermStore", new ExpressionValueResolver((s, v) => {

                    if (v != null)
                    {
                        var tgs = new TermGroupsSerializer();
                        var termGroupsExpressions = tgs.GetTermGroupSerializationExpressions();

                        var baseNamespace = PnPSerializationScope.Current?.BaseSchemaNamespace;
                        var termGroupType = Type.GetType($"{baseNamespace}.TermGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}", true);

                        var sourceSequence = s as ProvisioningSequence;

                        return(PnPObjectsMapper.MapObjects(sourceSequence.TermStore.TermGroups,
                            new CollectionFromModelToSchemaTypeResolver(termGroupType), 
                            termGroupsExpressions, 
                            true));
                    }
                    else
                    {
                        return (null);
                    }
                }));

                // Handle SiteCollections and hierarchycal subsites
                var siteCollectionTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteCollection, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var siteCollectionType = Type.GetType(siteCollectionTypeName, true);
                var subSiteTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Site, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var subSiteType = Type.GetType(subSiteTypeName, true);

                expressions.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Sequence.SiteCollections",
                    new SiteCollectionsAndSitesFromModelToSchemaTypeResolver(siteCollectionType));
                expressions.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteCollection.Sites",
                    new SiteCollectionsAndSitesFromModelToSchemaTypeResolver(subSiteType));
                expressions.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Site.Sites",
                    new SiteCollectionsAndSitesFromModelToSchemaTypeResolver(subSiteType));

                expressions.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteCollection.Templates", new ExpressionValueResolver((s, v) => {
                    return ConvertTemplateListToReferences(v);
                }));

                expressions.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.Site.Templates", new ExpressionValueResolver((s, v) =>
                {
                    return ConvertTemplateListToReferences(v);
                }));

                persistence.GetPublicInstanceProperty("Sequence")
                    .SetValue(
                        persistence,
                        PnPObjectsMapper.MapObjects(template.ParentHierarchy.Sequences,
                            new CollectionFromModelToSchemaTypeResolver(sequenceType), 
                            expressions, 
                            recursive: true));
            }
        }

        private static object ConvertTemplateListToReferences(object v)
        {
            var templateReferenceTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ProvisioningTemplateReference, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var templateReferenceType = Type.GetType(templateReferenceTypeName, true);

            var resultType = templateReferenceType.MakeArrayType();
            var resultArray = (Array)Activator.CreateInstance(resultType, ((IList)v).Count);
            var i = 0;

            foreach (var id in (IEnumerable)v)
            {
                var t = Activator.CreateInstance(templateReferenceType);
                t.SetPublicInstancePropertyValue("ID", id);
                resultArray.SetValue(t, i++);
            }

            return (resultArray.Length > 0 ? resultArray : null);
        }
    }
}
