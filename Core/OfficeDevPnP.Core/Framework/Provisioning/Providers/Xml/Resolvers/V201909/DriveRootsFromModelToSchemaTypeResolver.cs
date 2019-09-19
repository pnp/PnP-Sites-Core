using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.Resolvers
{
    /// <summary>
    /// Type resolver for DriveRoots from Model to Schema
    /// </summary>
    internal class DriveRootsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;

            // Get the RootFolder property
            var driveRoot = source as Model.Drive.DriveRoot;
            if (null != driveRoot)
            {
                var rootFolder = driveRoot.RootFolder;

                if (null != rootFolder)
                {
                    // Prepare the target types
                    var driveFolderTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFolder, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var driveFolderType = Type.GetType(driveFolderTypeName, true);
                    var driveFileTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFile, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                    var driveFileType = Type.GetType(driveFileTypeName, true);

                    // If we have folders or files
                    if ((rootFolder.DriveFolders != null &&
                        rootFolder.DriveFolders.Count > 0) ||
                        (rootFolder.DriveFiles != null &&
                        rootFolder.DriveFiles.Count > 0))
                    {
                        int itemsCount = (rootFolder.DriveFolders?.Count ?? 0) + (rootFolder.DriveFiles?.Count ?? 0);
                        var resultingItems = new Object[itemsCount];
                        var index = 0;

                        if (rootFolder.DriveFolders != null)
                        {
                            foreach (var df in rootFolder.DriveFolders)
                            {
                                var targetItem = Activator.CreateInstance(driveFolderType);
                                PnPObjectsMapper.MapProperties(df, targetItem, resolvers, recursive);
                                resultingItems.SetValue(targetItem, index);
                                index++;
                            }
                        }

                        if (rootFolder.DriveFiles != null)
                        {
                            foreach (var df in rootFolder.DriveFiles)
                            {
                                var targetItem = Activator.CreateInstance(driveFileType);
                                PnPObjectsMapper.MapProperties(df, targetItem, resolvers, recursive);
                                resultingItems.SetValue(targetItem, index);
                                index++;
                            }
                        }

                        result = resultingItems;
                    }
                }
            }

            return (result);
        }
    }
}
