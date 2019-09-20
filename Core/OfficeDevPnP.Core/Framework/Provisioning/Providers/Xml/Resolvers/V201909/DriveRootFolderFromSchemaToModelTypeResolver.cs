using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Drive;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Resolves the Drive Items from the Schema to the Model
    /// </summary>
    internal class DriveRootFolderFromSchemaToModelTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;

        public bool CustomCollectionResolver => false;

        public object Resolve(object source, Dictionary<String, IResolver> resolvers = null, Boolean recursive = false)
        {
            var result = new DriveRootFolder();

            var driveItems = source.GetPublicInstancePropertyValue("DriveItems") ?? source.GetPublicInstancePropertyValue("Items");

            var driveFolderTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFolder, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var driveFolderType = Type.GetType(driveFolderTypeName, true);
            var driveFileTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFile, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var driveFileType = Type.GetType(driveFileTypeName, true);

            if (null != driveItems)
            {
                foreach (var d in ((IEnumerable)driveItems))
                {
                    if (driveFolderType.IsInstanceOfType(d))
                    {
                        // If the item is a Folder
                        var targetItem = new DriveFolder();
                        PnPObjectsMapper.MapProperties(d, targetItem, resolvers, recursive);
                        result.DriveFolders.Add(targetItem);
                    }
                    else if (driveFileType.IsInstanceOfType(d))
                    {
                        // Else if the item is a File
                        var targetItem = new DriveFile();
                        PnPObjectsMapper.MapProperties(d, targetItem, resolvers, recursive);
                        result.DriveFiles.Add(targetItem);
                    }
                }
            }

            return (result);
        }
    }
}
