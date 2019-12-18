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
    /// Type resolver for DriveFolders and DriveFiles from Model to Schema
    /// </summary>
    internal class DriveItemsFromModelToSchemaTypeResolver : ITypeResolver
    {
        public string Name => this.GetType().Name;
        public bool CustomCollectionResolver => false;


        public object Resolve(object source, Dictionary<string, IResolver> resolvers = null, bool recursive = false)
        {
            Object result = null;
            Model.Drive.DriveFolderBase folder = null;

            // Let's see if the source is a DriveRoot object
            var driveRoot = source as Model.Drive.DriveRoot;
            if (driveRoot != null)
            {
                folder = driveRoot.RootFolder;
            }
            else
            {
                folder = (Model.Drive.DriveFolderBase)source;
            }

            if (null != folder)
            {
                // Prepare the target types
                var driveFolderTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFolder, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var driveFolderType = Type.GetType(driveFolderTypeName, true);
                var driveFileTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.DriveFile, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
                var driveFileType = Type.GetType(driveFileTypeName, true);

                // If we have folders or files
                if ((folder.DriveFolders != null &&
                    folder.DriveFolders.Count > 0) ||
                    (folder.DriveFiles != null &&
                    folder.DriveFiles.Count > 0))
                {
                    int itemsCount = (folder.DriveFolders?.Count ?? 0) + (folder.DriveFiles?.Count ?? 0);
                    var resultingItems = new Object[itemsCount];
                    var index = 0;

                    if (folder.DriveFolders != null)
                    {
                        foreach (var df in folder.DriveFolders)
                        {
                            var targetItem = Activator.CreateInstance(driveFolderType);
                            PnPObjectsMapper.MapProperties(df, targetItem, resolvers, recursive);
                            resultingItems.SetValue(targetItem, index);
                            index++;
                        }
                    }

                    if (folder.DriveFiles != null)
                    {
                        foreach (var df in folder.DriveFiles)
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

            //    var rootFolder = driveRoot.RootFolder;

            //    if (null != rootFolder)
            //    {

            //        // If we have folders or files
            //        if ((rootFolder.DriveFolders != null &&
            //            rootFolder.DriveFolders.Count > 0) ||
            //            (rootFolder.DriveFiles != null &&
            //            rootFolder.DriveFiles.Count > 0))
            //        {
            //            int itemsCount = (rootFolder.DriveFolders?.Count ?? 0) + (rootFolder.DriveFiles?.Count ?? 0);
            //            var resultingItems = new Object[itemsCount];
            //            var index = 0;

            //            if (rootFolder.DriveFolders != null)
            //            {
            //                foreach (var df in rootFolder.DriveFolders)
            //                {
            //                    var targetItem = Activator.CreateInstance(driveFolderType);
            //                    PnPObjectsMapper.MapProperties(df, targetItem, resolvers, recursive);
            //                    resultingItems.SetValue(targetItem, index);
            //                    index++;
            //                }
            //            }

            //            if (rootFolder.DriveFiles != null)
            //            {
            //                foreach (var df in rootFolder.DriveFiles)
            //                {
            //                    var targetItem = Activator.CreateInstance(driveFileType);
            //                    PnPObjectsMapper.MapProperties(df, targetItem, resolvers, recursive);
            //                    resultingItems.SetValue(targetItem, index);
            //                    index++;
            //                }
            //            }

            //            result = resultingItems;
            //        }
            //    }
            //}

            return (result);
        }
    }
}
