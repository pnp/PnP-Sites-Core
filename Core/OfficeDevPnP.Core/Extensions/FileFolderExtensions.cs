using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that holds the file and folder methods
    /// </summary>
    public static partial class FileFolderExtensions
    {
        /// <summary>
        /// Approves a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to approve</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static void ApproveFile(this Web web, string serverRelativeUrl, string comment)
        {
#if ONPREMISES
            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
#else
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
#endif
            web.Context.Load(file, x => x.Exists, x => x.CheckOutType);
            web.Context.ExecuteQueryRetry();

            if (file.Exists)
            {
                file.Approve(comment);
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Checks in a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkin</param>
        /// <param name="checkinType">The type of the checkin</param>
        /// <param name="comment">Message to be recorded with the approval</param>
        public static void CheckInFile(this Web web, string serverRelativeUrl, CheckinType checkinType, string comment)
        {
#if ONPREMISES

            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
#else
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
#endif
            var scope = new ConditionalScope(web.Context, () => file.ServerObjectIsNull.Value != true && file.Exists && file.CheckOutType != CheckOutType.None);

            using (scope.StartScope())
            {
                web.Context.Load(file);
            }
            web.Context.ExecuteQueryRetry();

            if (scope.TestResult.Value)
            {
                file.CheckIn(comment, checkinType);
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Checks out a file
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file to checkout</param>
        public static void CheckOutFile(this Web web, string serverRelativeUrl)
        {
#if ONPREMISES
            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
#else
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
#endif

            var scope = new ConditionalScope(web.Context, () => file.ServerObjectIsNull.Value != true && file.Exists && file.CheckOutType == CheckOutType.None);

            using (scope.StartScope())
            {
                web.Context.Load(file);
            }
            web.Context.ExecuteQueryRetry();

            if (scope.TestResult.Value)
            {
                file.CheckOut();
                web.Context.ExecuteQueryRetry();
            }
        }

        private static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        /// <summary>
        /// Creates a new document set as a child of an existing folder, with the specified content type ID.
        /// </summary>
        /// <param name="folder">Folder of the document set</param>
        /// <param name="documentSetName">Name of the document set</param>
        /// <param name="contentTypeId">Content type of the document set</param>
        /// <returns>The created Folder representing the document set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <example>
        ///     var setContentType = list.BestMatchContentTypeId(BuiltInContentTypeId.DocumentSet);
        ///     var set1 = list.RootFolder.CreateDocumentSet("Set 1", setContentType);
        /// </example>
        /// </remarks>
        public static Folder CreateDocumentSet(this Folder folder, string documentSetName, ContentTypeId contentTypeId)
        {
            if (folder == null) { throw new ArgumentNullException(nameof(folder)); }
            if (documentSetName == null) { throw new ArgumentNullException(nameof(documentSetName)); }
            if (contentTypeId == null) { throw new ArgumentNullException(nameof(contentTypeId)); }

            if (documentSetName.ContainsInvalidUrlChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateDocumentSet_The_argument_must_be_a_single_document_set_name_and_cannot_contain_path_characters_, nameof(documentSetName));
            }

            Log.Info(Constants.LOGGING_SOURCE, CoreResources.FieldAndContentTypeExtensions_CreateDocumentSet, documentSetName);

            var result = DocumentSet.DocumentSet.Create(folder.Context, folder, documentSetName, contentTypeId);
            folder.Context.ExecuteQueryRetry();

            var fullUri = new Uri(result.Value);
            var serverRelativeUrl = fullUri.AbsolutePath;
            var documentSetFolder = folder.Folders.GetByUrl(serverRelativeUrl);

            return documentSetFolder;
        }

        /// <summary>
        /// Converts a folder with the given name as a child of the List RootFolder. 
        /// </summary>
        /// <param name="list">List in which the folder exists</param>
        /// <param name="folderName">Folder name to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder ConvertFolderToDocumentSet(this List list, string folderName)
        {
            var folder = list.RootFolder.ResolveSubFolder(folderName);
            if (folder == null) throw new ArgumentException(CoreResources.FileFolderExtensions_FolderMissing);

            return ConvertFolderToDocumentSetImplementation(list, folder);
        }

        /// <summary>
        /// Converts a folder with the given name as a child of the List RootFolder. 
        /// </summary>
        /// <param name="list">List in which the folder exists</param>
        /// <param name="folder">Folder to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder ConvertFolderToDocumentSet(this List list, Folder folder)
        {
            return ConvertFolderToDocumentSetImplementation(list, folder);
        }

        /// <summary>
        /// Internal implementation of the Folder conversion to Document set
        /// </summary>
        /// <param name="list">Library in which the folder exists</param>
        /// <param name="folder">Folder to convert</param>
        /// <returns>The newly converted Document Set, so that additional operations (such as setting properties) can be done.</returns>
        private static Folder ConvertFolderToDocumentSetImplementation(this List list, Folder folder)
        {
            list.EnsureProperties(l => l.ContentTypes.Include(c => c.StringId));
            folder.Context.Load(folder.ListItemAllFields, l => l["ContentTypeId"]);
            folder.Context.ExecuteQueryRetry();
            var listItem = folder.ListItemAllFields;

            // If already a document set, just return the folder
            if (listItem["ContentTypeId"].ToString() == BuiltInContentTypeId.Folder) return folder;
            listItem["ContentTypeId"] = BuiltInContentTypeId.DocumentSet;

            // Add missing properties            
            listItem["HTML_x0020_File_x0020_Type"] = "Sharepoint.DocumentSet";
            folder.Properties["docset_LastRefresh"] = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss");
            folder.Properties["vti_contenttypeorder"] = string.Join(",", list.ContentTypes.ToList().Where(c => c.StringId.StartsWith(BuiltInContentTypeId.Document + "00"))?.Select(c => c.StringId));

            listItem.Update();
            folder.Update();
            list.Context.ExecuteQueryRetry();

            //Refresh Folder, otherwise 'Version conflict' error might be thrown on changing properties
            folder = list.RootFolder.ResolveSubFolder(folder.Name);
            return folder;
        }

        /// <summary>
        /// Creates a folder with the given name as a child of the Web. 
        /// Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created Folder, so that additional operations (such as setting properties) can be done.</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder CreateFolder(this Web web, string folderName)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = web.Folders;
            var folder = CreateFolderImplementation(folderCollection, folderName);
            return folder;
        }

        /// <summary>
        /// Creates a folder with the given name.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// <example>
        ///     var folder = list.RootFolder.CreateFolder("new-folder");
        /// </example>
        /// </remarks>
        public static Folder CreateFolder(this Folder parentFolder, string folderName)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var folder = CreateFolderImplementation(folderCollection, folderName, parentFolder);
            return folder;
        }

        private static Folder CreateFolderImplementation(FolderCollection folderCollection, string folderName, Folder parentFolder = null, params Expression<Func<Folder, object>>[] expressions)
        {
            ClientContext context = null;
            if (parentFolder != null)
            {
                context = parentFolder.Context as ClientContext;
            }

            List parentList = null;

            if (parentFolder != null)
            {
                parentFolder.EnsureProperty(p => p.Properties);
                if (parentFolder.Properties.FieldValues.ContainsKey("vti_listname"))
                {
                    if (context != null)
                    {
                        Guid parentListId = Guid.Parse((String)parentFolder.Properties.FieldValues["vti_listname"]);
                        parentList = context.Web.Lists.GetById(parentListId);
                        context.Load(parentList, l => l.BaseType, l => l.Title);
                        context.ExecuteQueryRetry();
                    }
                }
            }

            if (parentList == null)
            {
                // Create folder for library or common URL path
                var newFolder = folderCollection.Add(folderName);
                if (expressions != null && expressions.Any())
                {
                    folderCollection.Context.Load(newFolder, expressions);
                }
                else
                {
                    folderCollection.Context.Load(newFolder);
                }
                folderCollection.Context.ExecuteQueryRetry();
                return newFolder;
            }
            else
            {
                // Create folder for generic list
                ListItemCreationInformation newFolderInfo = new ListItemCreationInformation();
                newFolderInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                newFolderInfo.LeafName = folderName;
                parentFolder.EnsureProperty(f => f.ServerRelativeUrl);
                newFolderInfo.FolderUrl = parentFolder.ServerRelativeUrl;
                ListItem newFolderItem = parentList.AddItem(newFolderInfo);
                newFolderItem["Title"] = folderName;
                newFolderItem.Update();
                context.ExecuteQueryRetry();

                // Get the newly created folder
                var newFolder = parentFolder.Folders.GetByUrl(folderName);
                // Ensure all properties are loaded (to be compatible with the previous implementation)
                if (expressions != null && expressions.Any())
                {
                    context.Load(newFolder, expressions);
                }
                else
                {
                    context.Load(newFolder);
                }
                context.ExecuteQueryRetry();
                return (newFolder);
            }
        }

        /// <summary>
        /// Checks if a specific folder exists
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeFolderUrl">Folder to check</param>
        /// <returns>Returns true if folder exists</returns>
        public static bool DoesFolderExists(this Web web, string serverRelativeFolderUrl)
        {
#if ONPREMISES
            Folder folder = web.GetFolderByServerRelativeUrl(serverRelativeFolderUrl);
#else
            Folder folder = web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeFolderUrl));
#endif

            web.Context.Load(folder);
            bool exists = false;

            try
            {
                web.Context.ExecuteQueryRetry();
                exists = true;
            }
            catch
            {
                return false;
            }

            return exists;
        }

        /// <summary>
        /// Ensure that the folder structure is created. This also ensures hierarchy of folders.
        /// </summary>
        /// <param name="web">Web to be processed - can be root web or sub site</param>
        /// <param name="parentFolder">Parent folder</param>
        /// <param name="folderPath">Folder path</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The folder structure</returns>
        public static Folder EnsureFolder(this Web web, Folder parentFolder, string folderPath, params Expression<Func<Folder, object>>[] expressions)
        {
            web.EnsureProperties(w => w.ServerRelativeUrl);
            parentFolder.EnsureProperties(f => f.ServerRelativeUrl);

            var parentWebRelativeUrl = parentFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length);
            var webRelativeUrl = parentWebRelativeUrl + (parentWebRelativeUrl.EndsWith("/") ? "" : "/") + folderPath;

            return web.EnsureFolderPath(webRelativeUrl, expressions: expressions);
        }

        /// <summary>
        /// Checks if the folder exists at the top level of the web site, and if it does not exist creates it.
        /// Note it is more common to create folders within an existing Folder, such as the RootFolder of a List.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolder(this Web web, string folderName, params Expression<Func<Folder, object>>[] expressions)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = web.Folders;
            var folder = EnsureFolderImplementation(folderCollection, folderName, expressions: expressions);
            return folder;
        }

        /// <summary>
        /// Checks if the subfolder exists, and if it does not exist creates it.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolder(this Folder parentFolder, string folderName, params Expression<Func<Folder, object>>[] expressions)
        {
            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var folder = EnsureFolderImplementation(folderCollection, folderName, parentFolder, expressions);
            return folder;
        }

        private static Folder EnsureFolderImplementation(FolderCollection folderCollection, string folderName, Folder parentFolder = null, params Expression<Func<Folder, object>>[] expressions)
        {
            Folder folder = null;
            if (expressions != null && expressions.Any())
            {
                folderCollection.Context.Load(folderCollection, fc => fc.IncludeWithDefaultProperties(expressions));
            }
            else
            {
                folderCollection.Context.Load(folderCollection);
            }
            folderCollection.Context.ExecuteQueryRetry();
            foreach (Folder existingFolder in folderCollection)
            {
                if (string.Equals(existingFolder.Name, folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    folder = existingFolder;
                    break;
                }
            }

            if (folder == null)
            {
                folder = CreateFolderImplementation(folderCollection, folderName, parentFolder, expressions);
            }

            return folder;
        }

        /// <summary>
        /// Check if a folder exists with the specified path (relative to the web), and if not creates it (inside a list if necessary)
        /// </summary>
        /// <param name="web">Web to check for the specified folder</param>
        /// <param name="webRelativeUrl">Path to the folder, relative to the web site</param>
        /// <param name="expressions">List of lambda expressions of properties to load when retrieving the object</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// If the specified path is inside an existing list, then the folder is created inside that list.
        /// </para>
        /// <para>
        /// Any existing folders are traversed, and then any remaining parts of the path are created as new folders.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolderPath(this Web web, string webRelativeUrl, params Expression<Func<Folder, object>>[] expressions)
        {
            if (webRelativeUrl == null) { throw new ArgumentNullException(nameof(webRelativeUrl)); }

            //Web root folder should be returned if webRelativeUrl is empty
            if (webRelativeUrl.Length != 0 && string.IsNullOrWhiteSpace(webRelativeUrl)) { throw new ArgumentException(CoreResources.FileFolderExtensions_EnsureFolderPath_Folder_URL_is_required_, nameof(webRelativeUrl)); }

            // Check if folder exists
            if (!web.IsPropertyAvailable("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }

            var folderServerRelativeUrl = UrlUtility.Combine(web.ServerRelativeUrl, webRelativeUrl, "/");

            // Check if folder is inside a list
            var listCollection = web.Lists;
            web.Context.Load(listCollection, lc => lc.Include(l => l.RootFolder));
            web.Context.ExecuteQueryRetry();

            List containingList = null;

            foreach (var list in listCollection)
            {
                if (folderServerRelativeUrl.StartsWith(UrlUtility.Combine(list.RootFolder.ServerRelativeUrl, "/"), StringComparison.InvariantCultureIgnoreCase))
                {
                    // Load fields from the list
                    containingList = list;
                    break;
                }
            }

            // Start either at the root of the list or web
            string locationType = null;
            string rootUrl = null;
            Folder currentFolder = null;
            if (containingList == null)
            {
                locationType = "Web";
                currentFolder = web.EnsureProperty(w => w.RootFolder);
            }
            else
            {
                locationType = "List";
                currentFolder = containingList.RootFolder;
            }

            currentFolder.EnsureProperty(f => f.ServerRelativeUrl);
            rootUrl = currentFolder.ServerRelativeUrl;

            // Get remaining parts of the path and split
            var folderRootRelativeUrl = folderServerRelativeUrl.Substring(currentFolder.ServerRelativeUrl.Length);
            var childFolderNames = folderRootRelativeUrl.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var currentCount = 0;

            foreach (var folderName in childFolderNames)
            {
                currentCount++;

                // Find next part of the path
                var folderCollection = currentFolder.Folders;
                folderCollection.Context.Load(folderCollection);
                folderCollection.Context.ExecuteQueryRetry();
                Folder nextFolder = null;
                foreach (Folder existingFolder in folderCollection)
                {
                    if (string.Equals(existingFolder.Name, System.Net.WebUtility.UrlDecode(folderName), StringComparison.InvariantCultureIgnoreCase))
                    {
                        nextFolder = existingFolder;
                        break;
                    }
                }

                // Or create it
                if (nextFolder == null)
                {
                    var createPath = string.Join("/", childFolderNames, 0, currentCount);
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_CreateFolder0Under12, createPath, locationType, rootUrl);
                    if (locationType == "List")
                    {
                        createPath = createPath.Substring(0, createPath.Length - folderName.Length).TrimEnd('/');
                        var listUrl =
                            containingList.EnsureProperty(f => f.RootFolder).EnsureProperty(r => r.ServerRelativeUrl);
                        ListItemCreationInformation newFolderInfo = new ListItemCreationInformation();
                        newFolderInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                        newFolderInfo.LeafName = folderName;
                        newFolderInfo.FolderUrl = UrlUtility.Combine(listUrl, createPath);
                        ListItem newFolderItem = containingList.AddItem(newFolderInfo);

                        var titleField = web.Context.LoadQuery(containingList.Fields.Where(f => f.Id == BuiltInFieldId.Title));
                        web.Context.ExecuteQueryRetry();
                        if (titleField.Any())
                        {
                            newFolderItem["Title"] = folderName;
                        }

                        newFolderItem.Update();
                        containingList.Context.Load(newFolderItem);
                        containingList.Context.ExecuteQueryRetry();
                        nextFolder = web.GetFolderByServerRelativeUrl(UrlUtility.Combine(listUrl, createPath, folderName));
                        containingList.Context.Load(nextFolder);
                        containingList.Context.ExecuteQueryRetry();
                    }
                    else
                    {
                        nextFolder = folderCollection.Add(folderName);
                        folderCollection.Context.Load(nextFolder);
                        folderCollection.Context.ExecuteQueryRetry();
                    }
                }

                currentFolder = nextFolder;
            }
            if (expressions != null && expressions.Any())
            {
                web.Context.Load(currentFolder, expressions);
                web.Context.ExecuteQueryRetry();
            }
            return currentFolder;
        }

        /// <summary>
        /// Finds files in the web. Can be slow.
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static List<File> FindFiles(this Web web, string match)
        {
            Folder rootFolder = web.RootFolder;
            match = WildcardToRegex(match);
            List<File> files = new List<File>();

            ParseFiles(rootFolder, match, web.Context as ClientContext, ref files);

            return files;
        }

        /// <summary>
        /// Find files in the list, Can be slow.
        /// </summary>
        /// <param name="list">The list to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static List<File> FindFiles(this List list, string match)
        {
            Folder rootFolder = list.EnsureProperty(l => l.RootFolder);

            match = WildcardToRegex(match);
            List<File> files = new List<File>();

            ParseFiles(rootFolder, match, list.Context as ClientContext, ref files);

            return files;
        }

        /// <summary>
        /// Find files in a specific folder
        /// </summary>
        /// <param name="folder">The folder to process</param>
        /// <param name="match">a wildcard pattern to match</param>
        /// <returns>A list with the found <see cref="Microsoft.SharePoint.Client.File"/> objects</returns>
        public static List<File> FindFiles(this Folder folder, string match)
        {
            match = WildcardToRegex(match);
            List<File> files = new List<File>();

            ParseFiles(folder, match, folder.Context as ClientContext, ref files);

            return files;
        }

        /// <summary>
        /// Checks if the folder exists at the top level of the web site.
        /// </summary>
        /// <param name="web">Web to check for the named folder</param>
        /// <param name="folderName">Folder name to retrieve</param>
        /// <returns>true if the folder exists; false otherwise</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static bool FolderExists(this Web web, string folderName)
        {
            var folderCollection = web.Folders;
            var exists = FolderExistsImplementation(folderCollection, folderName);
            return exists;
        }

        /// <summary>
        /// Checks if the subfolder exists.
        /// </summary>
        /// <param name="parentFolder">Parent folder to check for the named subfolder</param>
        /// <param name="folderName">Folder name to retrieve</param>
        /// <returns>true if the folder exists; false otherwise</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// </remarks>
        public static bool FolderExists(this Folder parentFolder, string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            var folderCollection = parentFolder.Folders;
            var exists = FolderExistsImplementation(folderCollection, folderName);
            return exists;
        }

        private static bool FolderExistsImplementation(FolderCollection folderCollection, string folderName)
        {
            if (folderCollection == null)
            {
                throw new ArgumentNullException(nameof(folderCollection));
            }

            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            if (folderName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_CreateFolder_The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_, nameof(folderName));
            }

            folderCollection.Context.Load(folderCollection);
            folderCollection.Context.ExecuteQueryRetry();
            foreach (Folder folder in folderCollection)
            {
                if (folder.Name.Equals(folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Returns a file as string
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <returns>The file contents as a string</returns>
        public static string GetFileAsString(this Web web, string serverRelativeUrl)
        {
            string returnString = string.Empty;

#if ONPREMISES
            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
#else
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
#endif

            web.Context.Load(file);
            web.Context.ExecuteQueryRetry();
            ClientResult<Stream> stream = file.OpenBinaryStream();
            web.Context.ExecuteQueryRetry();

            using (Stream memStream = new MemoryStream())
            {
                CopyStream(stream.Value, memStream);
                memStream.Position = 0;
                StreamReader reader = new StreamReader(memStream);
                returnString = reader.ReadToEnd();
            }

            return returnString;
        }

        private static void ParseFiles(Folder folder, string match, ClientContext context, ref List<File> foundFiles)
        {
            FileCollection files = folder.Files;
            context.Load(files, fs => fs.Include(f => f.ServerRelativeUrl, f => f.Name, f => f.Title, f => f.TimeCreated, f => f.TimeLastModified));
            context.Load(folder.Folders);
            context.ExecuteQueryRetry();

            foreach (File file in files)
            {
                if (Regex.IsMatch(file.Name, match, RegexOptions.IgnoreCase))
                {
                    foundFiles.Add(file);
                }
            }

            foreach (Folder subfolder in folder.Folders)
            {
                ParseFiles(subfolder, match, context, ref foundFiles);
            }
        }

        /// <summary>
        /// Publishes a file existing on a server URL
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="serverRelativeUrl">the server relative URL of the file to publish</param>
        /// <param name="comment">Comment recorded with the publish action</param>
        public static void PublishFile(this Web web, string serverRelativeUrl, string comment)
        {
#if ONPREMISES
            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
#else
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
#endif

            web.Context.Load(file, x => x.Exists, x => x.CheckOutType);
            web.Context.ExecuteQueryRetry();

            if (file.Exists)
            {
                file.Publish(comment);
                web.Context.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Gets a folder with a given name in a given <see cref="Microsoft.SharePoint.Client.Folder"/>
        /// </summary>
        /// <param name="folder"><see cref="Microsoft.SharePoint.Client.Folder"/> in which to search for</param>
        /// <param name="folderName">Name of the folder to search for</param>
        /// <returns>The found <see cref="Microsoft.SharePoint.Client.Folder"/> if available, null otherwise</returns>
        public static Folder ResolveSubFolder(this Folder folder, string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                throw new ArgumentNullException(nameof(folderName));
            }

            folder.Context.Load(folder);
            folder.Context.Load(folder.Folders);
            folder.Context.ExecuteQueryRetry();

            foreach (Folder subFolder in folder.Folders)
            {
                if (subFolder.Name.Equals(folderName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return subFolder;
                }
            }

            return null;
        }

        /// <summary>
        /// Saves a remote file to a local folder
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="serverRelativeUrl">The server relative URL to the file</param>
        /// <param name="localPath">The local folder</param>
        /// <param name="localFileName">The local filename. If null the filename of the file on the server will be used</param>
        /// <param name="fileExistsCallBack">Optional callback function allowing to provide feedback if the file should be overwritten if it exists. The function requests a bool as return value and the string input contains the name of the file that exists.</param>
        public static void SaveFileToLocal(this Web web, string serverRelativeUrl, string localPath, string localFileName = null, Func<string, bool> fileExistsCallBack = null)
        {
            var clientContext = web.Context as ClientContext;

#if ONPREMISES
            var file = web.GetFileByServerRelativeUrl(serverRelativeUrl);
#else
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
#endif

            clientContext.Load(file);
            clientContext.ExecuteQueryRetry();

            ClientResult<Stream> stream = file.OpenBinaryStream();
            clientContext.ExecuteQueryRetry();

            var fileOut = Path.Combine(localPath, !string.IsNullOrEmpty(localFileName) ? localFileName : file.Name);

            if (!System.IO.File.Exists(fileOut) || (fileExistsCallBack != null && fileExistsCallBack(fileOut)))
            {
                using (Stream fileStream = new FileStream(fileOut, FileMode.Create))
                {
                    CopyStream(stream.Value, fileStream);
                }
            }
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Name of the file</param>
        /// <param name="localFilePath">Location of the file to be uploaded.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static File UploadFile(this Folder folder, string fileName, string localFilePath, bool overwriteIfExists)
        {
            if (folder == null)
            {
                throw new ArgumentNullException(nameof(folder));
            }

            if (localFilePath == null)
            {
                throw new ArgumentNullException(nameof(localFilePath));
            }

            if (!System.IO.File.Exists(localFilePath))
            {
                throw new FileNotFoundException("Local file was not found.", localFilePath);
            }

            using (var stream = System.IO.File.OpenRead(localFilePath))
                return folder.UploadFile(fileName, stream, overwriteIfExists);
        }

        /// <summary>
        /// Uploads a file to the specified folder.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Location of the file to be uploaded.</param>
        /// <param name="stream">A stream object that represents the file.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static File UploadFile(this Folder folder, string fileName, Stream stream, bool overwriteIfExists)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_UploadFile_Destination_file_name_is_required_, nameof(fileName));
            }

            if (fileName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_UploadFile_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_, nameof(fileName));
            }

            // Create the file
            var newFileInfo = new FileCreationInformation()
            {
                ContentStream = stream,
                Url = fileName,
                Overwrite = overwriteIfExists
            };

            Log.Debug(Constants.LOGGING_SOURCE, "Creating file info with Url '{0}'", newFileInfo.Url);
            var file = folder.Files.Add(newFileInfo);
            folder.Context.Load(file);
            folder.Context.ExecuteQueryRetry();

            return file;
        }

        /// <summary>
        /// Uploads a file to the specified folder by saving the binary directly (via webdav).
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Name of the file</param>
        /// <param name="localFilePath">Location of the file to be uploaded.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static File UploadFileWebDav(this Folder folder, string fileName, string localFilePath, bool overwriteIfExists)
        {
            if (folder == null)
            {
                throw new ArgumentNullException(nameof(folder));
            }

            if (localFilePath == null)
            {
                throw new ArgumentNullException(nameof(localFilePath));
            }

            if (!System.IO.File.Exists(localFilePath))
            {
                throw new FileNotFoundException("Local file was not found.", localFilePath);
            }

            using (var stream = System.IO.File.OpenRead(localFilePath))
                return folder.UploadFileWebDav(fileName, stream, overwriteIfExists);
        }

        /// <summary>
        /// Uploads a file to the specified folder by saving the binary directly (via webdav).
        /// Note: this method does not work using app only token.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Location of the file to be uploaded.</param>
        /// <param name="stream">A stream object that represents the file.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static File UploadFileWebDav(this Folder folder, string fileName, Stream stream, bool overwriteIfExists)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_UploadFile_Destination_file_name_is_required_, nameof(fileName));
            }

            if (fileName.ContainsInvalidFileFolderChars())
            {
                throw new ArgumentException(CoreResources.FileFolderExtensions_UploadFileWebDav_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_, nameof(fileName));
            }

            var serverRelativeUrl = UrlUtility.Combine(folder.ServerRelativeUrl, fileName);

            // Create uploadContext to get a proper ClientContext instead of a ClientRuntimeContext
            using (var uploadContext = folder.Context.Clone(folder.Context.Url))
            {
                Log.Debug(Constants.LOGGING_SOURCE, "Save binary direct (via webdav) to '{0}'", serverRelativeUrl);
                File.SaveBinaryDirect(uploadContext, serverRelativeUrl, stream, overwriteIfExists);
                uploadContext.ExecuteQueryRetry();
            }

            var file = folder.Files.GetByUrl(serverRelativeUrl);
            folder.Context.Load(file);
            folder.Context.ExecuteQueryRetry();

            return file;
        }

        /// <summary>
        /// Gets a file in a document library.
        /// </summary>
        /// <param name="folder">Folder containing the target file.</param>
        /// <param name="fileName">File name.</param>
        /// <returns>The target file if found, null if no file is found.</returns>
        public static File GetFile(this Folder folder, string fileName)
        {
            if (folder == null)
            {
                throw new ArgumentNullException(nameof(folder));
            }

            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException(nameof(fileName));
            }

            try
            {
                folder.EnsureProperties(f => f.ServerRelativeUrl);

                var fileServerRelativeUrl = UrlUtility.Combine(folder.ServerRelativeUrl, fileName);
                var context = folder.Context as ClientContext;

                var web = context.Web;

#if ONPREMISES
                var file = web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
#else
                var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(fileServerRelativeUrl));
#endif

                web.Context.Load(file);
                web.Context.ExecuteQueryRetry();
                return file;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    return null;
                }
                throw;
            }
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localFile">File to validate against.</param>
        public static bool VerifyIfUploadRequired(this File serverFile, string localFile)
        {
            if (localFile == null)
            {
                throw new ArgumentNullException(nameof(localFile));
            }

            if (!System.IO.File.Exists(localFile))
            {
                throw new FileNotFoundException("Local file was not found.", localFile);
            }

            using (var file = System.IO.File.OpenRead(localFile))
                return serverFile.VerifyIfUploadRequired(file);
        }

        /// <summary>
        /// Used to compare the server file to the local file.
        /// This enables users with faster download speeds but slow upload speeds to evaluate if the server file should be overwritten.
        /// </summary>
        /// <param name="serverFile">File located on the server.</param>
        /// <param name="localStream">Stream to validate against.</param>
        /// <returns></returns>
        public static bool VerifyIfUploadRequired(this File serverFile, Stream localStream)
        {
            if (serverFile == null)
            {
                throw new ArgumentNullException(nameof(serverFile));
            }

            if (localStream == null)
            {
                throw new ArgumentNullException(nameof(localStream));
            }

            byte[] serverHash = null;
            var streamResult = serverFile.OpenBinaryStream();
            serverFile.Context.ExecuteQueryRetry();

            // Hash contents
            HashAlgorithm ha = HashAlgorithm.Create();
            using (var serverStream = streamResult.Value)
            {
                serverHash = ha.ComputeHash(serverStream);
                //Console.WriteLine("Server hash: {0}", BitConverter.ToString(serverHash));
            }

            // Check hash (& rewind)
            var localHash = ha.ComputeHash(localStream);
            localStream.Position = 0;
            //Console.WriteLine("Local hash: {0}", BitConverter.ToString(localHash));

            // Compare hash
            var contentsMatch = true;
            for (var index = 0; index < serverHash.Length; index++)
            {
                if (serverHash[index] != localHash[index])
                {
                    //Console.WriteLine("Hash does not match");
                    contentsMatch = false;
                    break;
                }
            }

            localStream.Position = 0;
            return !contentsMatch;
        }

        /// <summary>
        /// Sets file properties using a dictionary.
        /// </summary>
        /// <param name="file">Target file object.</param>
        /// <param name="properties">Dictionary of properties to set.</param>
        /// <param name="checkoutIfRequired">Check out the file if necessary to set properties.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static void SetFileProperties(this File file, IDictionary<string, string> properties, bool checkoutIfRequired = true)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            if (properties == null)
            {
                throw new ArgumentNullException(nameof(properties));
            }

            var changedProperties = new Dictionary<string, string>();
            var changedPropertiesString = new StringBuilder();
            var context = file.Context;

            if (properties != null && properties.Count > 0)
            {
                // Get a reference to the target list, if any
                // and load file item properties
                var parentList = file.ListItemAllFields.ParentList;
                context.Load(parentList, l => l.ForceCheckout);
                context.Load(file.ListItemAllFields);
                context.Load(file.ListItemAllFields.FieldValuesAsText);
                try
                {
                    context.ExecuteQueryRetry();
                }
                catch (ServerException ex)
                {
                    // If this throws ServerException (does not belong to list), then shouldn't be trying to set properties)
                    // Handling the exception stating the "The object specified does not belong to a list."
                    if (ex.ServerErrorCode != -2146232832)
                    {
                        throw;
                    }
                }

                // Loop through and detect changes first, then, check out if required and apply
                foreach (var kvp in properties)
                {
                    var propertyName = kvp.Key;
                    var propertyValue = kvp.Value;

                    var fieldValues = file.ListItemAllFields.FieldValues;
                    var currentValue = string.Empty;
                    if (file.ListItemAllFields.FieldValues.ContainsKey(propertyName))
                    {
                        currentValue = file.ListItemAllFields.FieldValuesAsText[propertyName];
                    }

                    //LoggingUtility.Internal.TraceVerbose("*** Comparing property '{0}' to current '{1}', new '{2}'", propertyName, currentValue, propertyValue);
                    switch (propertyName.ToUpperInvariant())
                    {
                        case "CONTENTTYPE":
                            {
                                if (!currentValue.Equals(propertyValue, StringComparison.InvariantCultureIgnoreCase) && parentList != null)
                                {
                                    ContentType targetCT = parentList.GetContentTypeByName(propertyValue);
                                    context.ExecuteQueryRetry();

                                    if (targetCT != null)
                                    {
                                        changedProperties["ContentTypeId"] = targetCT.StringId;
                                        changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                    }
                                    else
                                    {
                                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_SetFileProperties_Error, propertyValue);
                                    }
                                }
                                break;
                            }
                        case "CONTENTTYPEID":
                            {
                                if (!currentValue.Equals(propertyValue, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                /*
                                var currentBase = currentValue.Substring(0, currentValue.Length - 34);
                                var sameValue = (currentBase == propertyValue);
                                if (!sameValue && propertyValue.Length >= 32 + 6 && propertyValue.Substring(propertyValue.Length - 34, 2) == "00")
                                {
                                    var propertyBase = propertyValue.Substring(0, propertyValue.Length - 34);
                                    sameValue = (currentBase == propertyBase);
                                }

                                if (!sameValue)
                                {
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                */
                                break;
                            }
                        case "PUBLISHINGASSOCIATEDCONTENTTYPE":
                            {
                                var testValue = ";#" + currentValue.Replace(", ", ";#") + ";#";
                                if (testValue != propertyValue)
                                {
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                break;
                            }
                        default:
                            {
                                if (currentValue != propertyValue)
                                {
                                    //Console.WriteLine("Setting property '{0}' to '{1}'", propertyName, propertyValue);
                                    changedProperties[propertyName] = propertyValue;
                                    changedPropertiesString.AppendFormat("{0}='{1}'; ", propertyName, propertyValue);
                                }
                                break;
                            }
                    }
                }

                if (changedProperties.Count > 0)
                {
                    Log.Info(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_UpdateFile0Properties1, file.Name, changedPropertiesString);
                    var checkOutRequired = false;

                    if (parentList != null)
                    {
                        checkOutRequired = parentList.ForceCheckout;
                    }

                    if (checkoutIfRequired && checkOutRequired && file.CheckOutType == CheckOutType.None)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Checking out file '{0}'", file.Name);
                        file.CheckOut();
                        context.ExecuteQueryRetry();
                    }

                    Log.Debug(Constants.LOGGING_SOURCE, "Set properties: {0}", file.Name);
                    foreach (var kvp in changedProperties)
                    {
                        var propertyName = kvp.Key;
                        var propertyValue = kvp.Value;

                        Log.Debug(Constants.LOGGING_SOURCE, " {0}={1}", propertyName, propertyValue);
                        file.ListItemAllFields[propertyName] = propertyValue;
                    }
                    file.ListItemAllFields.Update();
                    context.ExecuteQueryRetry();
                }
            }
        }

        /// <summary>
        /// Publishes a file based on the type of versioning required on the parent library.
        /// </summary>
        /// <param name="file">Target file to publish.</param>
        /// <param name="level">Target publish direction (Draft and Published only apply, Checkout is ignored).</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static void PublishFileToLevel(this File file, FileLevel level)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            var publishingRequired = false;
            var approvalRequired = false;

            if (level == FileLevel.Draft || level == FileLevel.Published)
            {
                var context = file.Context;

                bool normalFile = true;
                // Ensure that ListItemAllFields.ServerObjectIsNull is loaded
                try
                {
                    file.EnsureProperties(f => f.ListItemAllFields, f => f.CheckOutType, f => f.Name);
                }
                catch
                {
                    // Catch all errors...there's a valid scenario for this failing when this is not a file associated to a listitem
                    normalFile = false;
                }

                // Only access ListItemAllFields if the above load succeeded. If it didn't, accessing it will throw it back in the context, and the next
                // ExecuteQueryRetry will throw a 'The object specified does not belong to a list.' error.
                normalFile = normalFile && (!file.ListItemAllFields.ServerObjectIsNull ?? false); //normal files have listItemAllFields;
                var checkOutRequired = false;
                if (normalFile)
                {
                    var parentList = file.ListItemAllFields.ParentList;
                    context.Load(parentList,
                                l => l.EnableMinorVersions,
                                l => l.EnableModeration,
                                l => l.ForceCheckout);

                    try
                    {
                        context.ExecuteQueryRetry();
                        checkOutRequired = parentList.ForceCheckout;
                        publishingRequired = parentList.EnableMinorVersions; // minor versions implies that the file must be published
                        approvalRequired = parentList.EnableModeration;
                    }
                    catch (ServerException ex)
                    {
                        // Handling the exception stating the "The object specified does not belong to a list."
                        if (ex.ServerErrorCode != -2146232832)
                        {
                            // TODO Replace this with an errorcode as well, does not work with localized o365 tenants
                            if (ex.Message.StartsWith("Cannot invoke method or retrieve property from null object. Object returned by the following call stack is null.") &&
                                ex.Message.Contains("ListItemAllFields"))
                            {
                                // E.g. custom display form aspx page being uploaded to the libraries Forms folder
                                normalFile = false;
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }
                }

                if (file.CheckOutType != CheckOutType.None || checkOutRequired)
                {
                    Log.Debug(Constants.LOGGING_SOURCE, "Checking in file '{0}'", file.Name);
                    file.CheckIn("Checked in by provisioning", publishingRequired ? CheckinType.MinorCheckIn : CheckinType.MajorCheckIn);
                    context.ExecuteQueryRetry();
                }

                if (level == FileLevel.Published)
                {
                    if (publishingRequired)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Publishing file '{0}'", file.Name);
                        file.Publish("Published by provisioning");
                        context.ExecuteQueryRetry();
                    }

                    if (approvalRequired)
                    {
                        Log.Debug(Constants.LOGGING_SOURCE, "Approving file '{0}'", file.Name);
                        file.Approve("Approved by provisioning");
                        context.ExecuteQueryRetry();
                    }
                }
            }
        }

        private static string WildcardToRegex(string pattern)
        {
            return "^" + Regex.Escape(pattern).
                               Replace(@"\*", ".*").
                               Replace(@"\?", ".") + "$";
        }

    }
}
