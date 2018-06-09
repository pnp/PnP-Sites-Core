using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using File = Microsoft.SharePoint.Client.File;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;
using System;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectFiles : ObjectHandlerBase
    {
        private readonly string[] WriteableReadOnlyFields = new string[] { "publishingpagelayout", "contenttypeid" };

        // See https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f?ui=en-US&rs=en-US&ad=US
        public static readonly string[] BlockedExtensionsInNoScript = new string[] { ".asmx", ".ascx", ".aspx", ".htc", ".jar", ".master", ".swf", ".xap", ".xsf" };
        public static readonly string[] BlockedLibrariesInNoScript = new string[] { "_catalogs/theme", "style library", "_catalogs/lt", "_catalogs/wp" };

        public override string Name
        {
            get { return "Files"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Check if this is not a noscript site as we're not allowed to write to the web property bag is that one
                bool isNoScriptSite = web.IsNoScriptSite();
                web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);

                // Build on the fly the list of additional files coming from the Directories
                var directoryFiles = new List<Model.File>();
                foreach (var directory in template.Directories)
                {
                    var metadataProperties = directory.GetMetadataProperties();
                    directoryFiles.AddRange(directory.GetDirectoryFiles(metadataProperties));
                }

                var filesToProcess = template.Files.Union(directoryFiles).ToArray();
                var currentFileIndex = 0;
                var originalWeb = web; // Used to store and re-store context in case files are deployed to masterpage gallery
                foreach (var file in filesToProcess)
                {
                    file.Src = parser.ParseString(file.Src);
                    var targetFileName = !String.IsNullOrEmpty(file.TargetFileName) ? file.TargetFileName : file.Src;

                    currentFileIndex++;
                    WriteMessage($"File|{targetFileName}|{currentFileIndex}|{filesToProcess.Length}", ProvisioningMessageType.Progress);
                    var folderName = parser.ParseString(file.Folder);

                    if (folderName.ToLower().Contains("/_catalogs/"))
                    {
                        // Edge case where you have files in the template which should be provisioned to the site collection
                        // master page gallery and not to a connected subsite
                        web = web.Context.GetSiteCollectionContext().Web;
                        web.EnsureProperties(w => w.ServerRelativeUrl, w => w.Url);
                    }

                    if (folderName.ToLower().StartsWith((web.ServerRelativeUrl.ToLower())))
                    {
                        folderName = folderName.Substring(web.ServerRelativeUrl.Length);
                    }

                    if (SkipFile(isNoScriptSite, targetFileName, folderName))
                    {
                        // add log message
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Files_SkipFileUpload, targetFileName, folderName);
                        continue;
                    }

                    var folder = web.EnsureFolderPath(folderName);

                    var checkedOut = false;

                    var targetFile = folder.GetFile(template.Connector.GetFilenamePart(targetFileName));

                    if (targetFile != null)
                    {
                        if (file.Overwrite)
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Uploading_and_overwriting_existing_file__0_, targetFileName);
                            checkedOut = CheckOutIfNeeded(web, targetFile);

                            using (var stream = GetFileStream(template, file))
                            {
                                targetFile = UploadFile(template, file, folder, stream);
                            }
                        }
                        else
                        {
                            checkedOut = CheckOutIfNeeded(web, targetFile);
                        }
                    }
                    else
                    {
                        using (var stream = GetFileStream(template, file))
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Uploading_file__0_, targetFileName);
                            targetFile = UploadFile(template, file, folder, stream);
                        }

                        checkedOut = CheckOutIfNeeded(web, targetFile);
                    }

                    if (targetFile != null)
                    {
                        // Add the fileuniqueid tokens
#if !SP2013
                        targetFile.EnsureProperties(p => p.UniqueId, p => p.ServerRelativeUrl);
                        parser.AddToken(new FileUniqueIdToken(web, targetFile.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), targetFile.UniqueId));
                        parser.AddToken(new FileUniqueIdEncodedToken(web, targetFile.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart("/".ToCharArray()), targetFile.UniqueId));
#endif
                        if (file.Properties != null && file.Properties.Any())
                        {
                            Dictionary<string, string> transformedProperties = file.Properties.ToDictionary(property => property.Key, property => parser.ParseString(property.Value));
                            SetFileProperties(targetFile, transformedProperties, false);
                        }

#if !SP2013
                        bool webPartsNeedLocalization = false;
#endif
                        if (file.WebParts != null && file.WebParts.Any())
                        {
                            targetFile.EnsureProperties(f => f.ServerRelativeUrl);

                            var existingWebParts = web.GetWebParts(targetFile.ServerRelativeUrl).ToList();
                            foreach (var webPart in file.WebParts)
                            {
                                // check if the webpart is already set on the page
                                if (existingWebParts.FirstOrDefault(w => w.WebPart.Title == parser.ParseString(webPart.Title)) == null)
                                {
                                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Files_Adding_webpart___0___to_page, webPart.Title);
                                    var wpEntity = new WebPartEntity();
                                    wpEntity.WebPartTitle = parser.ParseString(webPart.Title);
                                    wpEntity.WebPartXml = parser.ParseXmlString(webPart.Contents).Trim(new[] { '\n', ' ' });
                                    wpEntity.WebPartZone = webPart.Zone;
                                    wpEntity.WebPartIndex = (int)webPart.Order;
                                    var wpd = web.AddWebPartToWebPartPage(targetFile.ServerRelativeUrl, wpEntity);
#if !SP2013
                                    if (webPart.Title.ContainsResourceToken())
                                    {
                                        // update data based on where it was added - needed in order to localize wp title
#if !SP2016
                                        wpd.EnsureProperties(w => w.ZoneId, w => w.WebPart, w => w.WebPart.Properties);
                                        webPart.Zone = wpd.ZoneId;
#else
                                        wpd.EnsureProperties(w => w.WebPart, w => w.WebPart.Properties);
#endif
                                        webPart.Order = (uint)wpd.WebPart.ZoneIndex;
                                        webPartsNeedLocalization = true;
                                    }
#endif
                                }
                            }
                        }

#if !SP2013
                        if (webPartsNeedLocalization)
                        {
                            file.LocalizeWebParts(web, parser, targetFile, scope);
                        }
#endif

                        switch (file.Level)
                        {
                            case Model.FileLevel.Published:
                                {
                                    targetFile.PublishFileToLevel(Microsoft.SharePoint.Client.FileLevel.Published);
                                    break;
                                }
                            case Model.FileLevel.Draft:
                                {
                                    targetFile.PublishFileToLevel(Microsoft.SharePoint.Client.FileLevel.Draft);
                                    break;
                                }
                            default:
                                {
                                    if (checkedOut)
                                    {
                                        targetFile.CheckIn("", CheckinType.MajorCheckIn);
                                        web.Context.ExecuteQueryRetry();
                                    }
                                    break;
                                }
                        }

                        // Don't set security when nothing is defined. This otherwise breaks on files set outside of a list
                        if (file.Security != null &&
                            (file.Security.ClearSubscopes == true || file.Security.CopyRoleAssignments == true || file.Security.RoleAssignments.Count > 0))
                        {
                            targetFile.ListItemAllFields.SetSecurity(parser, file.Security);
                        }
                    }

                    web = originalWeb; // restore context in case files are provisioned to the master page gallery #1059
                }
            }
            WriteMessage("Done processing files", ProvisioningMessageType.Completed);
            return parser;
        }

        private static bool CheckOutIfNeeded(Web web, File targetFile)
        {
            var checkedOut = false;
            try
            {
                web.Context.Load(targetFile, f => f.CheckOutType, f => f.CheckedOutByUser, f => f.ListItemAllFields.ParentList.ForceCheckout);
                web.Context.ExecuteQueryRetry();

                if (targetFile.ListItemAllFields.ServerObjectIsNull.HasValue
                    && !targetFile.ListItemAllFields.ServerObjectIsNull.Value
                    && targetFile.ListItemAllFields.ParentList.ForceCheckout)
                {
                    if (targetFile.CheckOutType == CheckOutType.None)
                    {
                        targetFile.CheckOut();
                    }
                    checkedOut = true;
                }
            }
            catch (ServerException ex)
            {
                // Handling the exception stating the "The object specified does not belong to a list."
                if (ex.ServerErrorCode != -2146232832)
                {
                    throw;
                }
            }
            return checkedOut;
        }


        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {

            return template;
        }


        public void SetFileProperties(File file, IDictionary<string, string> properties, bool checkoutIfRequired = true)
        {
            var context = file.Context;
            if (properties != null && properties.Count > 0)
            {
                // Get a reference to the target list, if any
                // and load file item properties
                var parentList = file.ListItemAllFields.ParentList;
                context.Load(parentList);
                context.Load(file.ListItemAllFields);
                try
                {
                    context.ExecuteQueryRetry();
                }
                catch (ServerException ex)
                {
                    // If this throws ServerException (does not belong to list), then shouldn't be trying to set properties)
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

                    var targetField = parentList.Fields.GetByInternalNameOrTitle(propertyName);
                    targetField.EnsureProperties(f => f.TypeAsString, f => f.ReadOnlyField);

                    // Changed by PaoloPia because there are fields like PublishingPageLayout
                    // which are marked as read-only, but have to be overwritten while uploading
                    // a publishing page file and which in reality can still be written
                    if (!targetField.ReadOnlyField || WriteableReadOnlyFields.Contains(propertyName.ToLower()))
                    {
                        switch (propertyName.ToUpperInvariant())
                        {
                            case "CONTENTTYPE":
                                {
                                    Microsoft.SharePoint.Client.ContentType targetCT = parentList.GetContentTypeByName(propertyValue);
                                    context.ExecuteQueryRetry();

                                    if (targetCT != null)
                                    {
                                        file.ListItemAllFields["ContentTypeId"] = targetCT.StringId;
                                    }
                                    else
                                    {
                                        Log.Error(Constants.LOGGING_SOURCE, "Content Type {0} does not exist in target list!", propertyValue);
                                    }
                                    break;
                                }
                            default:
                                {
                                    switch (targetField.TypeAsString)
                                    {
                                        case "User":
                                            var user = parentList.ParentWeb.EnsureUser(propertyValue);
                                            context.Load(user);
                                            context.ExecuteQueryRetry();

                                            if (user != null)
                                            {
                                                var userValue = new FieldUserValue
                                                {
                                                    LookupId = user.Id,
                                                };
                                                file.ListItemAllFields[propertyName] = userValue;
                                            }
                                            break;
                                        case "URL":
                                            var urlArray = propertyValue.Split(',');
                                            var linkValue = new FieldUrlValue();
                                            if (urlArray.Length == 2)
                                            {
                                                linkValue.Url = urlArray[0];
                                                linkValue.Description = urlArray[1];
                                            }
                                            else
                                            {
                                                linkValue.Url = urlArray[0];
                                                linkValue.Description = urlArray[0];
                                            }
                                            file.ListItemAllFields[propertyName] = linkValue;
                                            break;
                                        case "MultiChoice":
                                            var multiChoice = JsonUtility.Deserialize<String[]>(propertyValue);
                                            file.ListItemAllFields[propertyName] = multiChoice;
                                            break;
                                        case "LookupMulti":
                                            var lookupMultiValue = JsonUtility.Deserialize<FieldLookupValue[]>(propertyValue);
                                            file.ListItemAllFields[propertyName] = lookupMultiValue;
                                            break;
                                        case "TaxonomyFieldType":
                                            var taxonomyValue = JsonUtility.Deserialize<TaxonomyFieldValue>(propertyValue);
                                            file.ListItemAllFields[propertyName] = taxonomyValue;
                                            break;
                                        case "TaxonomyFieldTypeMulti":
                                            var taxonomyValueArray = JsonUtility.Deserialize<TaxonomyFieldValue[]>(propertyValue);
                                            file.ListItemAllFields[propertyName] = taxonomyValueArray;
                                            break;
                                        default:
                                            file.ListItemAllFields[propertyName] = propertyValue;
                                            break;
                                    }
                                    break;
                                }
                        }
                    }
                    file.ListItemAllFields.Update();
                    context.ExecuteQueryRetry();
                }
            }
        }

        /// <summary>
        /// Checks if a given file can be uploaded. Sites using NoScript can't handle all uploads
        /// </summary>
        /// <param name="isNoScriptSite">Is this a noscript site?</param>
        /// <param name="fileName">Filename to verify</param>
        /// <param name="folderName">Folder (library) to verify</param>
        /// <returns>True is the file will not be uploaded, false otherwise</returns>
        public static bool SkipFile(bool isNoScriptSite, string fileName, string folderName)
        {
            string fileExtionsion = Path.GetExtension(fileName).ToLower();
            if (isNoScriptSite)
            {
                if (!String.IsNullOrEmpty(fileExtionsion) && BlockedExtensionsInNoScript.Contains(fileExtionsion))
                {
                    // We need to skip this file
                    return true;
                }

                if (!String.IsNullOrEmpty(folderName))
                {
                    foreach (string blockedlibrary in BlockedLibrariesInNoScript)
                    {
                        if (folderName.ToLower().StartsWith(blockedlibrary))
                        {
                            // Can't write to this library, let's skip
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.Files.Any() | template.Directories.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }

        private static File UploadFile(ProvisioningTemplate template, Model.File file, Microsoft.SharePoint.Client.Folder folder, Stream stream)
        {
            File targetFile = null;
            var fileName = !String.IsNullOrEmpty(file.TargetFileName) ? file.TargetFileName : template.Connector.GetFilenamePart(file.Src);
            try
            {
                targetFile = folder.UploadFile(fileName, stream, file.Overwrite);
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorCode != -2130575306) //Error code: -2130575306 = The file is already checked out.
                {
                    //The file name might contain encoded characters that prevent upload. Decode it and try again.
                    fileName = WebUtility.UrlDecode(fileName);
                    try
                    {
                        targetFile = folder.UploadFile(fileName, stream, file.Overwrite);
                    }
                    catch (Exception)
                    {
                        //unable to Upload file, just ignore
                    }
                }
            }
            return targetFile;
        }

        /// <summary>
        /// Retrieves <see cref="Stream"/> from connector. If the file name contains special characters (e.g. "%20") and cannot be retrieved, a workaround will be performed
        /// </summary>
        private static Stream GetFileStream(ProvisioningTemplate template, Model.File file)
        {
            var fileName = file.Src;
            var container = String.Empty;
            if (fileName.Contains(@"\") || fileName.Contains(@"/"))
            {
                var tempFileName = fileName.Replace(@"/", @"\");
                container = fileName.Substring(0, tempFileName.LastIndexOf(@"\"));
                fileName = fileName.Substring(tempFileName.LastIndexOf(@"\") + 1);
            }

            // add the default provided container (if any)
            if (!String.IsNullOrEmpty(container))
            {
                if (!String.IsNullOrEmpty(template.Connector.GetContainer()))
                {
                    if (container.StartsWith("/"))
                    {
                        container = container.TrimStart("/".ToCharArray());
                    }
#if !NETSTANDARD2_0
                    if (template.Connector.GetType() == typeof(Connectors.AzureStorageConnector))
                    {
                        if (template.Connector.GetContainer().EndsWith("/"))
                        {
                            container = $@"{template.Connector.GetContainer()}{container}";
                        }
                        else
                        {
                            container = $@"{template.Connector.GetContainer()}/{container}";
                        }
                    }
                    else
                    {
                        container = $@"{template.Connector.GetContainer()}\{container}";
                    }
#else
                    container = $@"{template.Connector.GetContainer()}\{container}";
#endif
                }
            }
            else
            {
                container = template.Connector.GetContainer();
            }

            var stream = template.Connector.GetFileStream(fileName, container);
            if (stream == null)
            {
                //Decode the URL and try again
                fileName = WebUtility.UrlDecode(fileName);
                container = WebUtility.UrlDecode(container);
                stream = template.Connector.GetFileStream(fileName, container);
            }

            return stream;
        }
    }

    internal static class DirectoryExtensions
    {
        internal static Dictionary<String, Dictionary<String, String>> GetMetadataProperties(this Model.Directory directory)
        {
            Dictionary<String, Dictionary<String, String>> result = null;

            if (!string.IsNullOrEmpty(directory.MetadataMappingFile))
            {
                var metadataPropertiesStream = directory.ParentTemplate.Connector.GetFileStream(directory.MetadataMappingFile);
                if (metadataPropertiesStream != null)
                {
                    using (var sr = new StreamReader(metadataPropertiesStream))
                    {
                        var metadataPropertiesString = sr.ReadToEnd();
                        result = JsonConvert.DeserializeObject<Dictionary<String, Dictionary<String, String>>>(metadataPropertiesString);
                    }
                }
            }

            return (result);
        }

        internal static List<Model.File> GetDirectoryFiles(this Model.Directory directory,
            Dictionary<String, Dictionary<String, String>> metadataProperties = null)
        {
            var result = new List<Model.File>();

            // If the connector has a container specified we need to take that in account to find the files we need
            string folderToGrabFilesFrom = directory.Src;
            if (!String.IsNullOrEmpty(directory.ParentTemplate.Connector.GetContainer()))
            {
                folderToGrabFilesFrom = directory.ParentTemplate.Connector.GetContainer() + @"\" + directory.Src;
            }

            var files = directory.ParentTemplate.Connector.GetFiles(folderToGrabFilesFrom);

            if (!String.IsNullOrEmpty(directory.IncludedExtensions) && directory.IncludedExtensions != "*.*")
            {
                var includedExtensions = directory.IncludedExtensions.Split(new [] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                files = files.Where(f => includedExtensions.Contains($"*{Path.GetExtension(f).ToLower()}")).ToList();
            }

            if (!String.IsNullOrEmpty(directory.ExcludedExtensions))
            {
                var excludedExtensions = directory.ExcludedExtensions.Split(new [] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                files = files.Where(f => !excludedExtensions.Contains($"*{Path.GetExtension(f).ToLower()}")).ToList();
            }

            result.AddRange(from file in files
                            select new Model.File(
                                directory.Src + @"\" + file,
                                directory.Folder,
                                directory.Overwrite,
                                null, // No WebPartPages are supported with this technique
                                metadataProperties != null ? metadataProperties[directory.Src + @"\" + file] : null,
                                directory.Security,
                                directory.Level
                                ));

            if (directory.Recursive)
            {
                var subFolders = directory.ParentTemplate.Connector.GetFolders(folderToGrabFilesFrom);
                var parentFolder = directory;
                foreach (var folder in subFolders)
                {
                    directory.Src = parentFolder.Src + @"\" + folder;
                    directory.Folder = parentFolder.Folder + @"\" + folder;

                    result.AddRange(directory.GetDirectoryFiles(metadataProperties));

                    //Remove the subfolder path(added above) as the second subfolder should come under its parent folder and not under its sibling
                    parentFolder.Src = parentFolder.Src.Substring(0, parentFolder.Src.LastIndexOf(@"\"));
                    parentFolder.Folder = parentFolder.Folder.Substring(0, parentFolder.Folder.LastIndexOf(@"\"));
                }
            }

            return (result);
        }
    }
}