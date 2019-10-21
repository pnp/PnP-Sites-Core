using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers;
using OfficeDevPnP.Core.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Drive;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Serializers.V201909
{
    /// <summary>
    /// Class to serialize/deserialize the AAD settings
    /// </summary>
    [TemplateSchemaSerializer(
        MinimalSupportedSchemaVersion = XMLPnPSchemaVersion.V201909,
        SerializationSequence = 250, DeserializationSequence = 250,
        Scope = SerializerScope.Tenant)]
    internal class DriveSerializer : PnPBaseSchemaSerializer<Drive>
    {
        public override void Deserialize(object persistence, ProvisioningTemplate template)
        {
            var drive = persistence.GetPublicInstancePropertyValue("Drive");

            if (drive != null)
            {
                var expressions = new Dictionary<Expression<Func<Drive, Object>>, IResolver>();

                // Manage the DriveRoot items
                expressions.Add(d => d.DriveRoots, new DriveRootsFromSchemaToModelTypeResolver());
                expressions.Add(d => d.DriveRoots[0].RootFolder, new DriveRootFolderFromSchemaToModelTypeResolver());
                expressions.Add(d => d.DriveRoots[0].RootFolder.DriveFolders, 
                    new DriveItemsFromSchemaToModelTypeResolver(typeof(Model.Drive.DriveFolder)));
                expressions.Add(d => d.DriveRoots[0].RootFolder.DriveFiles, 
                    new DriveItemsFromSchemaToModelTypeResolver(typeof(Model.Drive.DriveFile)));

                PnPObjectsMapper.MapProperties(drive, template.ParentHierarchy.Drive, expressions, true);
            }
        }

        public override void Serialize(ProvisioningTemplate template, object persistence)
        {
            if (template.ParentHierarchy?.Drive?.DriveRoots != null)
            {
                var driveRootTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveRoot, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var driveRootType = Type.GetType(driveRootTypeName, false);
                var driveFolderTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFolder, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var driveFolderType = Type.GetType(driveFolderTypeName, false);

                if (driveRootType != null)
                {
                    var resolvers = new Dictionary<String, IResolver>();

                    //// Handle DriveRoot objects
                    //resolvers.Add($"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveRoot",
                    //    new DriveRootFolderFromModelToSchemaTypeResolver());

                    resolvers.Add($"{driveRootType}.DriveItems",
                        new DriveItemsFromModelToSchemaTypeResolver()); // DriveRootsFromModelToSchemaTypeResolver());
                    resolvers.Add($"{driveFolderType}.Items",
                        new DriveItemsFromModelToSchemaTypeResolver());


                    persistence.GetPublicInstanceProperty("Drive")
                        .SetValue(
                            persistence,
                            PnPObjectsMapper.MapObjects(template.ParentHierarchy?.Drive?.DriveRoots,
                                new CollectionFromModelToSchemaTypeResolver(driveRootType), 
                                resolvers, 
                                recursive: true));
                }
            }
        }
    }
}
