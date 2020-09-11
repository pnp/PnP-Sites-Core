using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using LanguageTemplateHash = System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Utilities;

namespace Microsoft.SharePoint.Client
{

    /// <summary>
    /// Class that deals with branding features
    /// </summary>
    public static partial class BrandingExtensions
    {
        const string AvailablePageLayouts = "__PageLayouts";
        const string DefaultPageLayout = "__DefaultPageLayout";
        const string AvailableWebTemplates = "__WebTemplates";
        const string InheritWebTemplates = "__InheritWebTemplates";
        const string InheritMaster = "__InheritsMasterUrl";
        const string InheritCustomMaster = "__InheritsCustomMasterUrl";
        const string InheritTheme = "__InheritsThemedCssFolderUrl";
        const string Inherit = "__Inherit";
        const string CurrentLookName = "Current";
        const string CAML_QUERY_FIND_BY_FILENAME = @"
                <View>
                    <Query>                
                        <Where>
                            <Eq>
                                <FieldRef Name='Name' />
                                <Value Type='Text'>{0}</Value>
                            </Eq>
                        </Where>
                     </Query>
                </View>";

        /// <summary>
        /// Checks if a composed look exists.
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <param name="composedLookName">Name of the composed look</param>
        /// <returns>true if it exists; otherwise false</returns>
        public static bool ComposedLookExists(this Web web, string composedLookName)
        {
            var found = GetComposedLook(web, composedLookName);
            return (found != null);
        }

        /// <summary>
        /// Creates (or updates) a composed look in the web site; usually this is done in the root site of the collection.
        /// </summary>
        /// <param name="web">Web to create the composed look in</param>
        /// <param name="lookName">Name of the theme</param>
        /// <param name="paletteFileName">File name of the palette file in the theme catalog of the site collection; path component ignored.</param>
        /// <param name="fontFileName">File name of the font file in the theme catalog of the site collection; path component ignored.</param>
        /// <param name="backgroundFileName">File name of the background image file in the theme catalog of the site collection; path component ignored.</param>
        /// <param name="masterFileName">File name of the master page in the mastepage catalog of the web site; path component ignored.</param>
        /// <param name="displayOrder">Display order of the composed look</param>
        /// <param name="replaceContent">Replace composed look if it already exists (default true)</param>
        public static void CreateComposedLookByName(this Web web, string lookName, string paletteFileName, string fontFileName, string backgroundFileName, string masterFileName, int displayOrder = 1, bool replaceContent = true)
        {
            var paletteUrl = default(string);
            var fontUrl = default(string);
            var backgroundUrl = default(string);
            var masterUrl = default(string);

            web.EnsureProperties(w => w.ServerRelativeUrl);

            if (!string.IsNullOrEmpty(paletteFileName))
            {
                paletteUrl = UrlUtility.Combine(web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(paletteFileName)));
            }
            if (!string.IsNullOrEmpty(fontFileName))
            {
                fontUrl = UrlUtility.Combine(web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(fontFileName)));
            }
            if (!string.IsNullOrEmpty(backgroundFileName))
            {
                backgroundUrl = UrlUtility.Combine(web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(backgroundFileName)));
            }

            if (!string.IsNullOrEmpty(masterFileName))
            {
                masterUrl = UrlUtility.Combine(web.ServerRelativeUrl, string.Format(Constants.MASTERPAGE_DIRECTORY, Path.GetFileName(masterFileName)));
            }

            CreateComposedLookByUrl(web, lookName, paletteUrl, fontUrl, backgroundUrl, masterUrl, displayOrder, replaceContent);
        }

        /// <summary>
        /// Creates (or updates) a composed look in the web site; usually this is done in the root site of the collection.
        /// </summary>
        /// <param name="web">Web to create the composed look in</param>
        /// <param name="lookName">Name of the theme</param>
        /// <param name="paletteServerRelativeUrl">URL of the palette file, usually in the theme catalog of the site collection</param>
        /// <param name="fontServerRelativeUrl">URL of the font file, usually in the theme catalog of the site collection</param>
        /// <param name="backgroundServerRelativeUrl">URL of the background image file, usually in /_layouts/15/images</param>
        /// <param name="masterServerRelativeUrl">URL of the master page, usually in the masterpage catalog of the web site</param>
        /// <param name="displayOrder">Display order of the composed look</param>
        /// <param name="replaceContent">Replace composed look if it already exists (default true)</param>
        public static void CreateComposedLookByUrl(this Web web, string lookName, string paletteServerRelativeUrl, string fontServerRelativeUrl, string backgroundServerRelativeUrl, string masterServerRelativeUrl, int displayOrder = 1, bool replaceContent = true)
        {
            web.EnsureProperties(w => w.ServerRelativeUrl);

            var composedLooksList = web.GetCatalog((int)ListTemplateType.DesignCatalog);

            // Check for existing, by name
            CamlQuery query = new CamlQuery();
            query.ViewXml = string.Format(CAML_QUERY_FIND_BY_FILENAME, lookName);
            var existingCollection = composedLooksList.GetItems(query);
            web.Context.Load(existingCollection);
            web.Context.ExecuteQueryRetry();
            ListItem item = existingCollection.FirstOrDefault();

            if (item == null)
            {
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_CreateComposedLook, lookName, web.ServerRelativeUrl);
                ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                item = composedLooksList.AddItem(itemInfo);
                item["Name"] = lookName;
                item["Title"] = lookName;
            }
            else
            {
                if (!replaceContent)
                {
                    throw new Exception("Composed look already exists, replace contents needs to be specified.");
                }
                Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_UpdateComposedLook, lookName, web.ServerRelativeUrl);
            }

            if (!string.IsNullOrEmpty(paletteServerRelativeUrl))
            {
                item["ThemeUrl"] = paletteServerRelativeUrl;
            }
            if (!string.IsNullOrEmpty(fontServerRelativeUrl))
            {
                item["FontSchemeUrl"] = fontServerRelativeUrl;
            }
            if (!string.IsNullOrEmpty(backgroundServerRelativeUrl))
            {
                item["ImageUrl"] = backgroundServerRelativeUrl;
            }
            // we use seattle master if anything else is not set
            if (string.IsNullOrEmpty(masterServerRelativeUrl))
            {
                item["MasterPageUrl"] = UrlUtility.Combine(web.ServerRelativeUrl, Constants.MASTERPAGE_SEATTLE);
            }
            else
            {
                item["MasterPageUrl"] = masterServerRelativeUrl;
            }

            item["DisplayOrder"] = displayOrder;
            item.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Retrieves the named composed look, overrides with specified palette, font, background and master page, and then recursively sets the specified values.
        /// </summary>
        /// <param name="web">Web to apply composed look to</param>
        /// <param name="lookName">Name of the composed look to apply; null will apply the override values only</param>
        /// <param name="paletteServerRelativeUrl">Override palette file URL to use</param>
        /// <param name="fontServerRelativeUrl">Override font file URL to use</param>
        /// <param name="backgroundServerRelativeUrl">Override background image file URL to use</param>
        /// <param name="masterServerRelativeUrl">Override master page file URL to use</param>
        /// <param name="resetSubsitesToInherit">false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit</param>
        /// <param name="updateRootOnly">false to apply to subsites; true (default) to only apply to specified site</param>
        public static void SetComposedLookByUrl(this Web web, string lookName, string paletteServerRelativeUrl = null, string fontServerRelativeUrl = null, string backgroundServerRelativeUrl = null, string masterServerRelativeUrl = null, bool resetSubsitesToInherit = false, bool updateRootOnly = true)
        {
            var paletteUrl = default(string);
            var fontUrl = default(string);
            var backgroundUrl = default(string);
            var masterUrl = default(string);

            if (!string.IsNullOrWhiteSpace(lookName))
            {
                var composedLooksList = web.GetCatalog((int)ListTemplateType.DesignCatalog);

                // Check for existing, by name
                CamlQuery query = new CamlQuery();
                query.ViewXml = string.Format(CAML_QUERY_FIND_BY_FILENAME, lookName);
                var existingCollection = composedLooksList.GetItems(query);
                web.Context.Load(existingCollection);
                web.Context.ExecuteQueryRetry();
                var item = existingCollection.FirstOrDefault();

                if (item != null)
                {
                    var lookPaletteUrl = item["ThemeUrl"] as FieldUrlValue;
                    if (lookPaletteUrl != null)
                    {
                        paletteUrl = new Uri(lookPaletteUrl.Url).AbsolutePath;
                    }
                    var lookFontUrl = item["FontSchemeUrl"] as FieldUrlValue;
                    if (lookFontUrl != null)
                    {
                        fontUrl = new Uri(lookFontUrl.Url).AbsolutePath;
                    }
                    var lookBackgroundUrl = item["ImageUrl"] as FieldUrlValue;
                    if (lookBackgroundUrl != null)
                    {
                        backgroundUrl = new Uri(lookBackgroundUrl.Url).AbsolutePath;
                    }
                    var lookMasterUrl = item["MasterPageUrl"] as FieldUrlValue;
                    if (lookMasterUrl != null)
                    {
                        masterUrl = new Uri(lookMasterUrl.Url).AbsolutePath;
                    }
                }
                else
                {
                    Log.Error(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_ComposedLookMissing, lookName);
                    throw new Exception($"Composed look '{lookName}' can not be found; pass null or empty to set look directly (not based on an existing entry)");
                }
            }

            if (!string.IsNullOrEmpty(paletteServerRelativeUrl))
            {
                paletteUrl = paletteServerRelativeUrl;
            }
            if (!string.IsNullOrEmpty(fontServerRelativeUrl))
            {
                fontUrl = fontServerRelativeUrl;
            }
            if (!string.IsNullOrEmpty(backgroundServerRelativeUrl))
            {
                backgroundUrl = backgroundServerRelativeUrl;
            }
            if (!string.IsNullOrEmpty(masterServerRelativeUrl))
            {
                masterUrl = masterServerRelativeUrl;
            }

            //URL decode retrieved URL's
            paletteUrl = System.Net.WebUtility.UrlDecode(paletteUrl);
            fontUrl = System.Net.WebUtility.UrlDecode(fontUrl);
            backgroundUrl = System.Net.WebUtility.UrlDecode(backgroundUrl);
            masterUrl = System.Net.WebUtility.UrlDecode(masterUrl);

            if (!string.IsNullOrEmpty(masterUrl))
            {
                web.SetMasterPageByUrl(masterUrl, resetSubsitesToInherit, updateRootOnly);
                web.SetCustomMasterPageByUrl(masterUrl, resetSubsitesToInherit, updateRootOnly);
            }

            if (!string.IsNullOrWhiteSpace(paletteUrl))
            {
                web.SetThemeByUrl(paletteUrl, fontUrl, backgroundUrl, resetSubsitesToInherit, updateRootOnly);
            }

            // Update/create the "Current" reference in the composed looks gallery
            string currentLookName = GetLocalizedCurrentValue(web);
            web.CreateComposedLookByUrl(currentLookName, paletteUrl, fontUrl, backgroundUrl, masterUrl, displayOrder: 0);
        }

        /// <summary>
        /// Recursively applies the specified palette, font, and background image.
        /// </summary>
        /// <param name="web">Web to apply to</param>
        /// <param name="paletteServerRelativeUrl">URL of palette file to apply</param>
        /// <param name="fontServerRelativeUrl">URL of font file to apply</param>
        /// <param name="backgroundServerRelativeUrl">URL of background image to apply</param>
        /// <param name="resetSubsitesToInherit">false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit</param>
        /// <param name="updateRootOnly">false (default) to apply to subsites; true to only apply to specified site</param>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static void SetThemeByUrl(this Web web, string paletteServerRelativeUrl, string fontServerRelativeUrl, string backgroundServerRelativeUrl, bool resetSubsitesToInherit = false, bool updateRootOnly = false)
        {
            var websToUpdate = new List<Web>();
            web.Context.Load(web, w => w.AllProperties, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_ApplyTheme, paletteServerRelativeUrl, web.ServerRelativeUrl);

            if(!web.IsNoScriptSite())
            {
                web.AllProperties[InheritTheme] = "False";
                web.Update();
            }
            web.ApplyTheme(paletteServerRelativeUrl, fontServerRelativeUrl, backgroundServerRelativeUrl, shareGenerated: true);
            web.Context.ExecuteQueryRetry();
            //web.Context.Load(web, w => w.ThemedCssFolderUrl);
            //var themedCssFolderUrl = childWeb.ThemedCssFolderUrl;
            websToUpdate.Add(web);

            if (!updateRootOnly)
            {
                var index = 0;
                while (index < websToUpdate.Count)
                {
                    var currentWeb = websToUpdate[index];
                    var websCollection = currentWeb.Webs;
                    web.Context.Load(websCollection, wc => wc.Include(w => w.AllProperties, w => w.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    foreach (var childWeb in websCollection)
                    {
                        var inheritThemeProperty = childWeb.GetPropertyBagValueString(InheritTheme, "");
                        bool inheritTheme = false;
                        if (!string.IsNullOrEmpty(inheritThemeProperty))
                        {
                            inheritTheme = string.Equals(childWeb.AllProperties[InheritTheme].ToString(), "True", StringComparison.InvariantCultureIgnoreCase);
                        }

                        if (resetSubsitesToInherit || inheritTheme)
                        {
                            Log.Debug(Constants.LOGGING_SOURCE, "Inherited: " + CoreResources.BrandingExtension_ApplyTheme, paletteServerRelativeUrl, childWeb.ServerRelativeUrl);
                            if (!web.IsNoScriptSite())
                            {
                                childWeb.AllProperties[InheritTheme] = "True";
                                //childWeb.ThemedCssFolderUrl = themedCssFolderUrl;
                                childWeb.Update();
                            }
                            // TODO: CSOM does not support the ThemedCssFolderUrl property yet (Nov 2014), so must call ApplyTheme at each level.
                            // This is very slow, so replace with simply setting the ThemedCssFolderUrl property instead once available.
                            childWeb.ApplyTheme(paletteServerRelativeUrl, fontServerRelativeUrl, backgroundServerRelativeUrl, shareGenerated: true);
                            web.Context.ExecuteQueryRetry();
                            websToUpdate.Add(childWeb);
                        }
                    }
                    index++;
                }
            }
        }


        /// <summary>
        /// Uploads the specified file (usually an spcolor or spfont file) to the web site themes gallery 
        /// (usually only exists in the root web of a site collection).
        /// </summary>
        /// <param name="web">Web site to upload to</param>
        /// <param name="localFilePath">Location of the file to be uploaded</param>
        /// <param name="themeFolderVersion">Leaf folder name to upload to; default is "15"</param>
        /// <returns>The uploaded file, with at least the ServerRelativeUrl property available</returns>
        public static File UploadThemeFile(this Web web, string localFilePath, string themeFolderVersion = "15")
        {
            if (localFilePath == null) { throw new ArgumentNullException(nameof(localFilePath)); }
            if (string.IsNullOrWhiteSpace(localFilePath)) { throw new ArgumentException(CoreResources.BrandingExtensions_UploadThemeFile_Source_file_path_is_required_, nameof(localFilePath)); }

            var fileName = Path.GetFileName(localFilePath);
            using (var localStream = new FileStream(localFilePath, FileMode.Open))
            {
                return UploadThemeFile(web, fileName, localStream, themeFolderVersion);
            }
        }

        /// <summary>
        /// Uploads the specified file (usually an spcolor or spfont file) to the web site themes gallery 
        /// (usually only exists in the root web of a site collection).
        /// </summary>
        /// <param name="web">Web site to upload to</param>
        /// <param name="fileName">Name of the file to upload</param>
        /// <param name="localFilePath">Location of the file to be uploaded</param>
        /// <param name="themeFolderVersion">Leaf folder name to upload to; default is "15"</param>
        /// <returns>The uploaded file, with at least the ServerRelativeUrl property available</returns>
        public static File UploadThemeFile(this Web web, string fileName, string localFilePath, string themeFolderVersion = "15")
        {
            if (fileName == null) { throw new ArgumentNullException(nameof(fileName)); }
            if (string.IsNullOrWhiteSpace(fileName)) { throw new ArgumentException(CoreResources.BrandingExtensions_UploadThemeFile_Destination_file_name_is_required_, nameof(fileName)); }
            if (localFilePath == null) { throw new ArgumentNullException(nameof(localFilePath)); }
            if (themeFolderVersion == null) throw new ArgumentNullException(nameof(themeFolderVersion));
            if (string.IsNullOrWhiteSpace(localFilePath)) { throw new ArgumentException(CoreResources.BrandingExtensions_UploadThemeFile_Source_file_path_is_required_, nameof(localFilePath)); }

            using (var localStream = new FileStream(localFilePath, FileMode.Open))
            {
                return UploadThemeFile(web, fileName, localStream, themeFolderVersion);
            }
        }

        /// <summary>
        /// Uploads the specified file (usually an spcolor or spfont file) to the web site themes gallery 
        /// (usually only exists in the root web of a site collection).
        /// </summary>
        /// <param name="web">Web site to upload to</param>
        /// <param name="fileName">Name of the file to upload</param>
        /// <param name="localStream">Stream containing the contents of the file</param>
        /// <param name="themeFolderVersion">Leaf folder name to upload to; default is "15"</param>
        /// <returns>The uploaded file, with at least the ServerRelativeUrl property available</returns>
        public static File UploadThemeFile(this Web web, string fileName, Stream localStream, string themeFolderVersion = "15")
        {
            if (fileName == null) { throw new ArgumentNullException(nameof(fileName)); }
            if (localStream == null) { throw new ArgumentNullException(nameof(localStream)); }
            if (string.IsNullOrWhiteSpace(fileName)) { throw new ArgumentException(CoreResources.BrandingExtensions_UploadThemeFile_Destination_file_name_is_required_, nameof(fileName)); }
            if (fileName.ContainsInvalidUrlChars())
            {
                throw new ArgumentException(CoreResources.BrandingExtensions_UploadThemeFile_The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_, nameof(fileName));
            }

            // Theme catalog only exists at site collection root
            var themesList = web.GetCatalog((int)ListTemplateType.ThemeCatalog);
            var themesFolder = themesList.RootFolder.EnsureFolder(themeFolderVersion);
            return themesFolder.UploadFile(fileName, localStream, true);
        }

        /// <summary>
        /// Can be used to deploy page layouts to master page gallery. 
        /// <remarks>Should be only used with root web of site collection where publishing features are enabled.</remarks>
        /// </summary>
        /// <param name="web">Web as the root site of the publishing site collection</param>
        /// <param name="sourceFilePath">Full path to the file which will be uploaded</param>
        /// <param name="title">Title for the page layout</param>
        /// <param name="description">Description for the page layout</param>
        /// <param name="associatedContentTypeID">Associated content type ID</param>
        /// <param name="folderHierarchy">Folder hierarchy where the page layouts will be deployed</param>
        public static void DeployPageLayout(this Web web, string sourceFilePath, string title, string description, string associatedContentTypeID, string folderHierarchy = "")
        {
            web.DeployMasterPageGalleryItem(sourceFilePath, title, description, associatedContentTypeID, Constants.PAGE_LAYOUT_CONTENT_TYPE, folderHierarchy);
        }

        /// <summary>
        /// Can be used to deploy html page layouts to master page gallery. 
        /// <remarks>Should be only used with root web of site collection where publishing features are enabled.</remarks>
        /// </summary>
        /// <param name="web">Web as the root site of the publishing site collection</param>
        /// <param name="sourceFilePath">Full path to the file which will be uploaded</param>
        /// <param name="title">Title for the page layout</param>
        /// <param name="description">Description for the page layout</param>
        /// <param name="associatedContentTypeID">Associated content type ID</param>
        /// <param name="folderHierarchy">Folder hierarchy where the html page layouts will be deployed</param>
        public static void DeployHtmlPageLayout(this Web web, string sourceFilePath, string title, string description, string associatedContentTypeID, string folderHierarchy = "")
        {
            web.DeployMasterPageGalleryItem(sourceFilePath, title, description, associatedContentTypeID, Constants.HTMLPAGE_LAYOUT_CONTENT_TYPE, folderHierarchy);
        }

        /// <summary>
        /// Private method to support all kinds of file uploads to the master page gallery
        /// </summary>
        /// <param name="web">Web as the root site of the publishing site collection</param>
        /// <param name="sourceFilePath">Full path to the file which will be uploaded</param>
        /// <param name="title">Title for the page layout</param>
        /// <param name="description">Description for the page layout</param>
        /// <param name="associatedContentTypeID">Associated content type ID</param>
        /// <param name="itemContentTypeId">Content type id for the item.</param>
        /// <param name="folderHierarchy">Folder hierarchy where the file will be uploaded</param>
        private static void DeployMasterPageGalleryItem(this Web web, string sourceFilePath, string title, string description, string associatedContentTypeID, string itemContentTypeId, string folderHierarchy = "")
        {
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                throw new ArgumentNullException(nameof(sourceFilePath));
            }

            if (!System.IO.File.Exists(sourceFilePath))
            {
                throw new FileNotFoundException("File for param sourceFilePath file does not exist", sourceFilePath);
            }

            string fileName = Path.GetFileName(sourceFilePath);
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_DeployPageLayout, fileName, web.Context.Url);

            // Get the path to the file which we are about to deploy
            List masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            Folder rootFolder = masterPageGallery.RootFolder;
            web.Context.Load(masterPageGallery);
            web.Context.Load(rootFolder);
            web.Context.ExecuteQueryRetry();

            // Create folder structure inside master page gallery, if does not exists
            // For e.g.: _catalogs/masterpage/contoso/
            web.EnsureFolder(rootFolder, folderHierarchy);

            var fileBytes = System.IO.File.ReadAllBytes(sourceFilePath);

            // Use CSOM to upload the file in
            FileCreationInformation newFile = new FileCreationInformation();
            newFile.Content = fileBytes;
            newFile.Url = UrlUtility.Combine(rootFolder.ServerRelativeUrl, folderHierarchy, fileName);
            newFile.Overwrite = true;

            File uploadFile = rootFolder.Files.Add(newFile);
            web.Context.Load(uploadFile);
            web.Context.ExecuteQueryRetry();

            // Check out the file if needed
            if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
            {
                if (uploadFile.CheckOutType == CheckOutType.None)
                {
                    uploadFile.CheckOut();
                }
            }

            // Get content type for ID to assign associated content type information
            ContentType associatedCt = web.GetContentTypeById(associatedContentTypeID);

            var listItem = uploadFile.ListItemAllFields;
            listItem["Title"] = title;
            listItem["MasterPageDescription"] = description;
            // set the item as page layout
            listItem["ContentTypeId"] = itemContentTypeId;
            // Set the associated content type ID property
            listItem["PublishingAssociatedContentType"] = $";#{associatedCt.Name};#{associatedCt.Id};#";
            listItem["UIVersion"] = Convert.ToString(15);
            listItem.Update();

            // Check in the page layout if needed
            if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
            {
                uploadFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                if (masterPageGallery.EnableModeration)
                {
                    listItem.File.Publish(string.Empty);
                }
            }
            web.Context.ExecuteQueryRetry();

        }

        /// <summary>
        /// Deploys a new masterpage
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="sourceFilePath">The path to the source file</param>
        /// <param name="title">The title of the masterpage</param>
        /// <param name="description">The description of the masterpage</param>
        /// <param name="uiVersion">UIVersion of the masterpage</param>
        /// <param name="defaultCSSFile">DefaultCSSFile of the masterpage</param>
        /// <param name="folderPath">FolderPath of the masterpage</param>
        public static File DeployMasterPage(this Web web, string sourceFilePath, string title, string description, string uiVersion = "15", string defaultCSSFile = "", string folderPath = "")
        {
            if (string.IsNullOrEmpty(sourceFilePath))
                throw new ArgumentNullException(nameof(sourceFilePath));

            if (!System.IO.File.Exists(sourceFilePath))
                throw new FileNotFoundException("File for param sourceFilePath not found.", sourceFilePath);

            string fileName = Path.GetFileName(sourceFilePath);
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_DeployMasterPage, fileName, web.Context.Url);

            // Get the path to the file which we are about to deploy
            List masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            Folder rootFolder = masterPageGallery.RootFolder;
            web.Context.Load(masterPageGallery);
            web.Context.Load(rootFolder);
            web.Context.ExecuteQueryRetry();

            // Create folder if does not exists
            if (!String.IsNullOrEmpty(folderPath))
            {
                web.EnsureFolder(rootFolder, folderPath);
            }

            // Get the file name from the provided path
            var fileBytes = System.IO.File.ReadAllBytes(sourceFilePath);

            // Use CSOM to upload the file in
            FileCreationInformation newFile = new FileCreationInformation();
            newFile.Content = fileBytes;
            newFile.Url = UrlUtility.Combine(rootFolder.ServerRelativeUrl, folderPath, fileName);
            newFile.Overwrite = true;

            File uploadFile = rootFolder.Files.Add(newFile);
            web.Context.Load(uploadFile);
            web.Context.ExecuteQueryRetry();


            var listItem = uploadFile.ListItemAllFields;
            if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
            {
                if (uploadFile.CheckOutType == CheckOutType.None)
                {
                    uploadFile.CheckOut();
                }
            }

            listItem["Title"] = title;
            listItem["MasterPageDescription"] = description;
            // Set content type as master page
            listItem["ContentTypeId"] = Constants.MASTERPAGE_CONTENT_TYPE;
            listItem["UIVersion"] = uiVersion;
            listItem.Update();
            if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
            {
                uploadFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                if (masterPageGallery.EnableModeration)
                {
                    listItem.File.Publish(string.Empty);
                }
            }
            web.Context.Load(listItem);
            web.Context.ExecuteQueryRetry();

            return uploadFile;
        }

        /// <summary>
        /// Can be used to set master page and custom master page in single command
        /// </summary>
        /// <param name="web">Current web</param>
        /// <param name="masterPageName">Name of the master page</param>
        /// <param name="customMasterPageName">Name of the custom master page</param>
        /// <exception cref="System.ArgumentException">Thrown when masterPageName or customMasterPageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when masterPageName or customMasterPageName is null</exception>
        public static void SetMasterPagesByName(this Web web, string masterPageName, string customMasterPageName)
        {
            if (string.IsNullOrEmpty(masterPageName))
            {
                throw (masterPageName == null)
                  ? new ArgumentNullException(nameof(masterPageName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(masterPageName));
            }
            if (string.IsNullOrEmpty(customMasterPageName))
            {
                throw (customMasterPageName == null)
                  ? new ArgumentNullException(nameof(customMasterPageName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(customMasterPageName));
            }

            web.SetMasterPageByName(masterPageName);
            web.SetCustomMasterPageByName(customMasterPageName);
        }


        /// <summary>
        /// Can be used to set master page and custom master page in single command
        /// </summary>
        /// <param name="web">Current web</param>
        /// <param name="masterPageUrl">Url of the master page</param>
        /// <param name="customMasterPageUrl">Url of the custom master page</param>
        /// <exception cref="System.ArgumentException">Thrown when masterPageName or customMasterPageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when masterPageName or customMasterPageName is null</exception>
        public static void SetMasterPagesByUrl(this Web web, string masterPageUrl, string customMasterPageUrl)
        {
            if (string.IsNullOrEmpty(masterPageUrl))
            {
                throw (masterPageUrl == null)
                  ? new ArgumentNullException(nameof(masterPageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(masterPageUrl));
            }
            if (string.IsNullOrEmpty(customMasterPageUrl))
            {
                throw (customMasterPageUrl == null)
                  ? new ArgumentNullException(nameof(customMasterPageUrl))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(customMasterPageUrl));
            }

            web.SetMasterPageByUrl(masterPageUrl);
            web.SetCustomMasterPageByUrl(customMasterPageUrl);
        }

        /// <summary>
        /// Master page is set by using master page name. Master page is set from the current web.
        /// </summary>
        /// <param name="web">Current web</param>
        /// <param name="masterPageName">Name of the master page. Path is resolved from this.</param>
        /// <exception cref="System.ArgumentException">Thrown when masterPageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when masterPageName is null</exception>  
        public static void SetMasterPageByName(this Web web, string masterPageName)
        {
            if (string.IsNullOrEmpty(masterPageName))
            {
                throw (masterPageName == null)
                  ? new ArgumentNullException(nameof(masterPageName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(masterPageName));
            }
            string masterPageUrl = GetRelativeUrlForMasterByName(web, masterPageName);
            if (!string.IsNullOrEmpty(masterPageUrl))
            {
                SetMasterPageByUrl(web, masterPageUrl);
            }
        }


        /// <summary>
        /// Master page is set by using master page name. Master page is set from the current web.
        /// </summary>
        /// <param name="web">Current web</param>
        /// <param name="masterPageName">Name of the master page. Path is resolved from this.</param>
        /// <exception cref="System.ArgumentException">Thrown when masterPageName is a zero-length string or contains only white space</exception>
        /// <exception cref="System.ArgumentNullException">Thrown when masterPageName is null</exception>  
        public static void SetCustomMasterPageByName(this Web web, string masterPageName)
        {
            if (string.IsNullOrEmpty(masterPageName))
            {
                throw (masterPageName == null)
                  ? new ArgumentNullException(nameof(masterPageName))
                  : new ArgumentException(CoreResources.Exception_Message_EmptyString_Arg, nameof(masterPageName));
            }

            string masterPageUrl = GetRelativeUrlForMasterByName(web, masterPageName);
            if (!string.IsNullOrEmpty(masterPageUrl))
            {
                SetCustomMasterPageByUrl(web, masterPageUrl);
            }
        }

        /// <summary>
        /// Returns the relative URL for a masterpage
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="masterPageName">The name of the masterpage, e.g. 'default' or 'seattle'</param>
        /// <returns></returns>
        [SuppressMessage("Microsoft.Globalization", "CA1308:NormalizeStringsToUppercase",
            Justification = "URLs are commonly standardised to lower case.")]
        public static string GetRelativeUrlForMasterByName(this Web web, string masterPageName)
        {
            if (string.IsNullOrEmpty(masterPageName))
                throw new ArgumentNullException(nameof(masterPageName));

            List masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            CamlQuery query = new CamlQuery();
            // Use query Scope='RecursiveAll' to iterate through sub folders of Master page library because we might have file in folder hierarchy
            query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Contains><FieldRef Name='FileRef'/><Value Type='Text'>.master</Value></Contains></Where></Query></View>";
            ListItemCollection galleryItems = masterPageGallery.GetItems(query);
            web.Context.Load(masterPageGallery);
            web.Context.Load(galleryItems);
            web.Context.ExecuteQueryRetry();
            foreach (var item in galleryItems)
            {
                var fileRef = item["FileRef"].ToString();
                if (fileRef.ToUpperInvariant().Contains(masterPageName.ToUpperInvariant()))
                {
                    return fileRef.ToLowerInvariant();
                }
            }
            return string.Empty;
        }

        private static string GetLocalizedCurrentValue(this Web web)
        {
            web.EnsureProperties(w => w.Language);
            ClientResult<string> currentTranslated = Utilities.Utility.GetLocalizedString(web.Context, "$Resources:Current", "core", (int)web.Language);
            web.Context.ExecuteQueryRetry();
            return currentTranslated.Value;
        }


        /// <summary>
        /// Returns the current theme of a web
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <returns>Entity with attributes of current composed look, or null if none</returns>
        public static ThemeEntity GetCurrentComposedLook(this Web web)
        {
            var themeName = GetLocalizedCurrentValue(web);
            return GetComposedLook(web, themeName);
        }

        /// <summary>
        /// Returns the named composed look from the web gallery
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <param name="composedLookName">Name of the composed look to retrieve</param>
        /// <returns>Entity with the attributes of the composed look, or null if the composed look does not exists or cannot be determined</returns>
        public static ThemeEntity GetComposedLook(this Web web, string composedLookName)
        {
            // List of OOB composed looks
            List<string> defaultComposedLooks = new List<string>(new[] { "Orange", "Sea Monster", "Green", "Lime", "Nature", "Blossom", "Sketch", "City", "Orbit", "Grey", "Characters", "Office", "Breeze", "Immerse", "Red", "Purple", "Wood" });

            // ThemeEntity object that will be 
            ThemeEntity theme = null;

            List designCatalog = web.GetCatalog((int)ListTemplateType.DesignCatalog);
            const string camlString = @"
            <View>  
                <Query> 
                </Query> 
                <OrderBy>
                   <FieldRef Name='Modified' />
                </OrderBy>
                <ViewFields>
                    <FieldRef Name='Name' />
                    <FieldRef Name='ImageUrl' />
                    <FieldRef Name='MasterPageUrl' />
                    <FieldRef Name='FontSchemeUrl' />
                    <FieldRef Name='ThemeUrl' />
                </ViewFields> 
            </View>";

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = camlString;

            ListItemCollection themes = designCatalog.GetItems(camlQuery);
            web.Context.Load(themes);
            web.Context.Load(web, w => w.Url);
            web.Context.ExecuteQueryRetry();

            var siteCollectionUrl = "";
            var subSitePath = "";
            using (var cc = web.Context.Clone(web.Url))
            {
                cc.Load(cc.Site, s => s.Url);
                cc.ExecuteQueryRetry();
                siteCollectionUrl = cc.Site.Url;
                subSitePath = web.Url.Replace(siteCollectionUrl, "");
            }

            string currentLookName = GetLocalizedCurrentValue(web);

            if (themes.Count > 0)
            {
                List<string> customComposedLooks = new List<string>();

                // Iterate over the existing composed looks to figure out the current composed look
                foreach (var themeItem in themes)
                {
                    string masterPageUrl = null;
                    string themeUrl = null;
                    string imageUrl = null;
                    string fontUrl = null;
                    string name = null;

                    if (themeItem["MasterPageUrl"] != null && themeItem["MasterPageUrl"].ToString().Length > 0)
                    {
                        masterPageUrl = (themeItem["MasterPageUrl"] as FieldUrlValue).Url;
                    }
                    if (themeItem["ImageUrl"] != null && themeItem["ImageUrl"].ToString().Length > 0)
                    {
                        imageUrl = (themeItem["ImageUrl"] as FieldUrlValue).Url;
                    }
                    if (themeItem["FontSchemeUrl"] != null && themeItem["FontSchemeUrl"].ToString().Length > 0)
                    {
                        fontUrl = (themeItem["FontSchemeUrl"] as FieldUrlValue).Url;
                    }
                    if (themeItem["ThemeUrl"] != null && themeItem["ThemeUrl"].ToString().Length > 0)
                    {
                        themeUrl = (themeItem["ThemeUrl"] as FieldUrlValue).Url;
                    }
                    if (themeItem["Name"] != null && themeItem["Name"].ToString().Length > 0)
                    {
                        name = themeItem["Name"] as String;
                    }

                    if (name != null)
                    {
                        if (!name.Equals(currentLookName, StringComparison.InvariantCultureIgnoreCase) &&
                            !defaultComposedLooks.Contains(name))
                        {
                            customComposedLooks.Add(name);
                        }

                        if (name.Equals(composedLookName, StringComparison.InvariantCultureIgnoreCase))
                        {
                            theme = new ThemeEntity();
                            if (themeItem["ThemeUrl"] != null && themeItem["ThemeUrl"].ToString().Length > 0)
                            {
                                theme.Theme = (themeItem["ThemeUrl"] as FieldUrlValue).Url;
                            }
                            if (themeItem["MasterPageUrl"] != null && themeItem["MasterPageUrl"].ToString().Length > 0)
                            {
                                theme.MasterPage = (themeItem["MasterPageUrl"] as FieldUrlValue).Url;
                            }
                            if (themeItem["FontSchemeUrl"] != null && themeItem["FontSchemeUrl"].ToString().Length > 0)
                            {
                                theme.Font = (themeItem["FontSchemeUrl"] as FieldUrlValue).Url;
                            }
                            if (themeItem["ImageUrl"] != null && themeItem["ImageUrl"].ToString().Length > 0)
                            {
                                theme.BackgroundImage = (themeItem["ImageUrl"] as FieldUrlValue).Url;
                            }
                        }
                    }
                }

                // return here if we did not find the requested theme...it does not exist.
                if (theme == null)
                {
                    return theme;
                }

                // For a brand new clean site everything is null. Once you apply another OOB composed look and then re-apply the default 
                // Office composed look the theme information will be populated.
                if (theme.BackgroundImage == null &&
                    theme.Font == null &&
                    theme.MasterPage == null &&
                    theme.Theme == null &&
                    web.IsUsingOfficeTheme())
                {
                    theme.Name = "Office";
                    theme.MasterPage = $"{subSitePath}/_catalogs/masterpage/seattle.master";
                    theme.Theme = "/_catalogs/theme/15/palette001.spcolor";
                    theme.IsCustomComposedLook = false;
                }
                else
                {
                    // Loop over the defined composed look and get the one that matches the information gathered from the "current" composed look
                    bool themeMatched = false;

                    //first loop avoids comparing with "current" entry in order to detect oob themes
                    //if no match then the second run includes "current"
                    for (var i = 0; i < 2; i++)
                    {
                        if (themeMatched)
                        {
                            break;
                        }

                        foreach (var themeItem in themes)
                        {
                            string masterPageUrl = null;
                            string themeUrl = null;
                            string imageUrl = null;
                            string fontUrl = null;
                            string name = "";

                            if (themeItem["MasterPageUrl"] != null && themeItem["MasterPageUrl"].ToString().Length > 0)
                            {
                                masterPageUrl = System.Net.WebUtility.UrlDecode((themeItem["MasterPageUrl"] as FieldUrlValue).Url);
                            }
                            if (themeItem["ImageUrl"] != null && themeItem["ImageUrl"].ToString().Length > 0)
                            {
                                imageUrl = System.Net.WebUtility.UrlDecode((themeItem["ImageUrl"] as FieldUrlValue).Url);
                            }
                            if (themeItem["FontSchemeUrl"] != null && themeItem["FontSchemeUrl"].ToString().Length > 0)
                            {
                                fontUrl = System.Net.WebUtility.UrlDecode((themeItem["FontSchemeUrl"] as FieldUrlValue).Url);
                            }
                            if (themeItem["ThemeUrl"] != null && themeItem["ThemeUrl"].ToString().Length > 0)
                            {
                                themeUrl = System.Net.WebUtility.UrlDecode((themeItem["ThemeUrl"] as FieldUrlValue).Url);
                            }
                            if (themeItem["Name"] != null && themeItem["Name"].ToString().Length > 0)
                            {
                                name = themeItem["Name"] as String;
                            }

                            // Exclude current from this comparison as otherwise we'll never detect the actual theme name
                            if (!name.Equals(currentLookName, StringComparison.InvariantCultureIgnoreCase) && i == 0)
                            {
                                // Note: do not take in account the ImageUrl field as this will point to a copied image in case of a sub site
                                if (IsMatchingTheme(theme, masterPageUrl, themeUrl, fontUrl))
                                {
                                    theme.Name = name;
                                    theme.IsCustomComposedLook = !defaultComposedLooks.Contains(theme.Name);

                                    // Restore the default composed look image URL
                                    if (imageUrl != null)
                                    {
                                        theme.BackgroundImage = imageUrl;
                                    }

                                    // We're taking the first matching composed look
                                    themeMatched = true;
                                    break;
                                }
                            }
                        }
                    }

                    // special case, theme files have been deployed via api and when applying the proper theme the "current" was not set
                    if (!string.IsNullOrEmpty(theme.Name) && theme.Name.Equals(currentLookName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!web.IsUsingOfficeTheme())
                        {
                            // Assume the the last added custom theme is what the site is using
                            for (var i = themes.Count; i-- > 0;)
                            {
                                var themeItem = themes[i];
                                if (themeItem["Name"] != null && customComposedLooks.Contains(themeItem["Name"] as string))
                                {
                                    if (themeItem["ThemeUrl"] != null && themeItem["ThemeUrl"].ToString().Length > 0)
                                    {
                                        theme.Theme = (themeItem["ThemeUrl"] as FieldUrlValue).Url;
                                    }
                                    if (themeItem["MasterPageUrl"] != null && themeItem["MasterPageUrl"].ToString().Length > 0)
                                    {
                                        theme.MasterPage = (themeItem["MasterPageUrl"] as FieldUrlValue).Url;
                                    }
                                    if (themeItem["FontSchemeUrl"] != null && themeItem["FontSchemeUrl"].ToString().Length > 0)
                                    {
                                        theme.Font = (themeItem["FontSchemeUrl"] as FieldUrlValue).Url;
                                    }
                                    if (themeItem["ImageUrl"] != null && themeItem["ImageUrl"].ToString().Length > 0)
                                    {
                                        theme.BackgroundImage = (themeItem["ImageUrl"] as FieldUrlValue).Url;
                                    }
                                    if (themeItem["Name"] != null && themeItem["Name"].ToString().Length > 0)
                                    {
                                        theme.Name = themeItem["Name"] as String;
                                    }

                                    theme.IsCustomComposedLook = true;
                                    break;
                                }
                            }
                        }
                    }

                }
            }

            if (theme == null)
            {
                return theme;
            }

            var designPreviewThemedCssFolderUrl = web.GetPropertyBagValueString("DesignPreviewThemedCssFolderUrl", null);

            // If name still is "Current" and there isn't a PreviewThemedCssFolderUrl 
            // property in the property bag then we can't correctly determine the set 
            // composed look...so return null
            if (!string.IsNullOrEmpty(theme.Name) && theme.Name.Equals(currentLookName, StringComparison.InvariantCultureIgnoreCase)
                && string.IsNullOrEmpty(designPreviewThemedCssFolderUrl))
            {
                return null;
            }

            // Clean up the fully qualified URLs
            if (theme.BackgroundImage != null && theme.BackgroundImage.IndexOf(siteCollectionUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                theme.BackgroundImage = theme.BackgroundImage.Replace(siteCollectionUrl, "");
            }
            if (theme.Theme != null && theme.Theme.IndexOf(siteCollectionUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                theme.Theme = theme.Theme.Replace(siteCollectionUrl, "");
            }
            if (theme.Font != null && theme.Font.IndexOf(siteCollectionUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                theme.Font = theme.Font.Replace(siteCollectionUrl, "");
            }
            if (theme.MasterPage != null && theme.MasterPage.IndexOf(siteCollectionUrl, StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                theme.MasterPage = theme.MasterPage.Replace(siteCollectionUrl, "");
            }

            return theme;
        }

        /// <summary>
        /// Compares master page URL, theme URL and font URL values to current theme entity to check if they are the same.
        /// Handles also possible null values. Point is to figure out which theme is the one that is currently
        /// being selected as "Current"
        /// </summary>
        /// <param name="theme">Current theme entity to compare values to</param>
        /// <param name="masterPageUrl">Master page URL</param>
        /// <param name="themeUrl">Theme URL</param>
        /// <param name="fontUrl">Font URL</param>
        /// <returns>Returns true if matches</returns>
        private static bool IsMatchingTheme(ThemeEntity theme, string masterPageUrl, string themeUrl, string fontUrl)
        {
            var themeUrlHasValue = false;
            var fontUrlHasValue = false;

            //Is Masterpage Url meaningful for compare?
            var masterPageUrlHasValue = !string.IsNullOrEmpty(theme.MasterPage);

            // Is theme URL meaningful for compare?
            if (!string.IsNullOrEmpty(theme.Theme))
            {
                themeUrlHasValue = true;
            }

            // Is font URL meaningful for compare?
            if (!string.IsNullOrEmpty(theme.Font))
            {
                fontUrlHasValue = true;
            }

            // Should we compare all of the values?
            if (masterPageUrlHasValue && themeUrlHasValue && fontUrlHasValue)
            {
                if (!string.IsNullOrEmpty(theme.MasterPage) && theme.MasterPage.Equals(masterPageUrl, StringComparison.InvariantCultureIgnoreCase) &&
                    theme.Theme.Equals(themeUrl, StringComparison.InvariantCultureIgnoreCase) &&
                    theme.Font.Equals(fontUrl, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }

            // Should we compare only master page and theme URL?
            if (masterPageUrlHasValue && themeUrlHasValue && !fontUrlHasValue)
            {
                if (!string.IsNullOrEmpty(theme.MasterPage) && theme.MasterPage.Equals(masterPageUrl, StringComparison.InvariantCultureIgnoreCase) &&
                    theme.Theme.Equals(themeUrl, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }

            // Should we compare only master page and font value?
            if (masterPageUrlHasValue && !themeUrlHasValue && fontUrlHasValue)
            {
                if (!string.IsNullOrEmpty(theme.MasterPage) && theme.MasterPage.Equals(masterPageUrl, StringComparison.InvariantCultureIgnoreCase) &&
                    theme.Font.Equals(fontUrl, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }

            // Should we only compare master page
            if (masterPageUrlHasValue && !themeUrlHasValue && !fontUrlHasValue)
            {
                if (!string.IsNullOrEmpty(theme.MasterPage) && theme.MasterPage.Equals(masterPageUrl, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsUsingOfficeTheme(this Web web)
        {
            try
            {
                var ti = web.ThemeInfo;
                web.Context.Load(ti);
                var accentText = ti.GetThemeShadeByName("AccentText");
                var backgroundOverlay = ti.GetThemeShadeByName("BackgroundOverlay");
                var bodyText = ti.GetThemeShadeByName("BodyText");
                web.Context.ExecuteQueryRetry();

                var accentTextRGB = accentText.Value.Substring(2);
                var backgroundOverlayARGB = backgroundOverlay.Value.Substring(2);
                var bodyTextRGB = bodyText.Value.Substring(2);

                if (accentTextRGB.Equals("0072C6") &&
                    backgroundOverlayARGB.Equals("FFFFFF") &&
                    bodyTextRGB.Equals("444444") &&
                    ti.ThemeBackgroundImageUri == null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                //Can cause issues in OnPrem-installations, because the theme is not found by SharePoint (see https://github.com/OfficeDev/PnP/issues/671).
                //Assuming the branding is not the office theme to continue.
                return false;
            }
        }


        /// <summary>
        /// Gets a page layout from the master page catalog. Can be called with paramter as "pagelayout.aspx" or as full path like "_catalog/masterpage/pagelayout.aspx"
        /// </summary>
        /// <param name="web">root web</param>
        /// <param name="pageLayoutName">name of the page layout to retrieve</param>
        /// <returns>ListItem holding the page layout, null if not found</returns>
        public static ListItem GetPageLayoutListItemByName(this Web web, string pageLayoutName)
        {
            if (string.IsNullOrEmpty(pageLayoutName))
            {
                throw new ArgumentNullException(nameof(pageLayoutName));
            }

            var pageLayoutFolder = Path.GetDirectoryName(pageLayoutName);
            var pageLayoutNameWithoutPath = Path.GetFileNameWithoutExtension(pageLayoutName);

            if (!string.IsNullOrEmpty(pageLayoutFolder))
            {
                // strip trailing /
                pageLayoutFolder = pageLayoutFolder.Replace("\\", "/");
                if (pageLayoutFolder.Substring(0, 1).Equals("/", StringComparison.InvariantCultureIgnoreCase))
                {
                    pageLayoutFolder = pageLayoutFolder.Substring(1);
                }

                pageLayoutNameWithoutPath = $"{pageLayoutFolder}/{pageLayoutNameWithoutPath}";
            }

            var masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            web.Context.Load(masterPageGallery, x => x.RootFolder.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            var fileRefValue = $"{masterPageGallery.RootFolder.ServerRelativeUrl}/{pageLayoutNameWithoutPath}.aspx";
            var query = new CamlQuery();
            // Use query Scope='RecursiveAll' to iterate through sub folders of Master page library because we might have file in folder hierarchy
            query.ViewXml =
                $"<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>{fileRefValue}</Value></Eq></Where></Query></View>";
            var galleryItems = masterPageGallery.GetItems(query);
            web.Context.Load(masterPageGallery);
            web.Context.Load(galleryItems);
            web.Context.ExecuteQueryRetry();
            return galleryItems.Count > 0 ? galleryItems[0] : null;
        }

        /// <summary>
        /// Set master page by using given URL as parameter. Suitable for example in cases where you want sub sites to reference root site master page gallery. This is typical with publishing sites.
        /// </summary>
        /// <param name="web">Context web</param>
        /// <param name="masterPageServerRelativeUrl">URL to the master page.</param>
        /// <param name="resetSubsitesToInherit">false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit</param>
        /// <param name="updateRootOnly">false (default) to apply to subsites; true to only apply to specified site</param>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static void SetMasterPageByUrl(this Web web, string masterPageServerRelativeUrl, bool resetSubsitesToInherit = false, bool updateRootOnly = false)
        {
            if (string.IsNullOrEmpty(masterPageServerRelativeUrl)) { throw new ArgumentNullException(nameof(masterPageServerRelativeUrl)); }

            var websToUpdate = new List<Web>();
            web.Context.Load(web, w => w.AllProperties, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_SetMasterUrl, masterPageServerRelativeUrl, web.ServerRelativeUrl);
            web.AllProperties[InheritMaster] = "False";
            web.MasterUrl = masterPageServerRelativeUrl;
            web.Update();
            web.Context.ExecuteQueryRetry();
            websToUpdate.Add(web);

            if (!updateRootOnly)
            {
                var index = 0;
                while (index < websToUpdate.Count)
                {
                    var currentWeb = websToUpdate[index];
                    var websCollection = currentWeb.Webs;
                    web.Context.Load(websCollection, wc => wc.Include(w => w.AllProperties, w => w.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    foreach (var childWeb in websCollection)
                    {
                        if (childWeb.GetBaseTemplateId() != "APP#0")
                        {
                            var inheritThemeProperty = childWeb.GetPropertyBagValueString(InheritTheme, "");
                            bool inheritTheme = false;
                            if (!string.IsNullOrEmpty(inheritThemeProperty))
                            {
                                inheritTheme = string.Equals(childWeb.AllProperties[InheritTheme].ToString(), "True", StringComparison.InvariantCultureIgnoreCase);
                            }

                            if (resetSubsitesToInherit || inheritTheme)
                            {
                                Log.Debug(Constants.LOGGING_SOURCE, "Inherited: " + CoreResources.BrandingExtension_SetMasterUrl, masterPageServerRelativeUrl, childWeb.ServerRelativeUrl);
                                childWeb.AllProperties[InheritMaster] = "True";
                                childWeb.MasterUrl = masterPageServerRelativeUrl;
                                childWeb.Update();
                                web.Context.ExecuteQueryRetry();
                                websToUpdate.Add(childWeb);
                            }
                        }
                    }
                    index++;
                }
            }
        }

        /// <summary>
        /// Set Custom master page by using given URL as parameter. Suitable for example in cases where you want sub sites to reference root site master page gallery. This is typical with publishing sites.
        /// </summary>
        /// <param name="web">Context web</param>
        /// <param name="masterPageServerRelativeUrl">URL to the master page.</param>
        /// <param name="resetSubsitesToInherit">false (default) to apply to currently inheriting subsites only; true to force all subsites to inherit</param>
        /// <param name="updateRootOnly">false (default) to apply to subsites; true to only apply to specified site</param>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        public static void SetCustomMasterPageByUrl(this Web web, string masterPageServerRelativeUrl, bool resetSubsitesToInherit = false, bool updateRootOnly = false)
        {
            if (string.IsNullOrEmpty(masterPageServerRelativeUrl)) { throw new ArgumentNullException(nameof(masterPageServerRelativeUrl)); }

            var websToUpdate = new List<Web>();
            web.Context.Load(web, w => w.AllProperties, w => w.ServerRelativeUrl);
            web.Context.ExecuteQueryRetry();

            Log.Info(Constants.LOGGING_SOURCE, CoreResources.BrandingExtension_SetCustomMasterUrl, masterPageServerRelativeUrl, web.ServerRelativeUrl);
            web.AllProperties[InheritCustomMaster] = "False";
            web.CustomMasterUrl = masterPageServerRelativeUrl;
            web.Update();
            web.Context.ExecuteQueryRetry();
            websToUpdate.Add(web);

            if (!updateRootOnly)
            {
                var index = 0;
                while (index < websToUpdate.Count)
                {
                    var currentWeb = websToUpdate[index];
                    var websCollection = currentWeb.Webs;
                    web.Context.Load(websCollection, wc => wc.Include(w => w.AllProperties, w => w.ServerRelativeUrl));
                    web.Context.ExecuteQueryRetry();
                    foreach (var childWeb in websCollection)
                    {
                        if (childWeb.GetBaseTemplateId() != "APP#0")
                        {
                            var inheritThemeProperty = childWeb.GetPropertyBagValueString(InheritTheme, "");
                            var inheritTheme = false;
                            if (!string.IsNullOrEmpty(inheritThemeProperty))
                            {
                                inheritTheme = string.Equals(childWeb.AllProperties[InheritTheme].ToString(), "True", StringComparison.InvariantCultureIgnoreCase);
                            }

                            if (resetSubsitesToInherit || inheritTheme)
                            {
                                Log.Debug(Constants.LOGGING_SOURCE, "Inherited: " + CoreResources.BrandingExtension_SetCustomMasterUrl, masterPageServerRelativeUrl, childWeb.ServerRelativeUrl);
                                childWeb.AllProperties[InheritCustomMaster] = "True";
                                childWeb.CustomMasterUrl = masterPageServerRelativeUrl;
                                childWeb.Update();
                                web.Context.ExecuteQueryRetry();
                                websToUpdate.Add(childWeb);
                            }
                        }
                    }
                    index++;
                }
            }

        }

        /// <summary>
        /// Sets specific page layout the default page layout for the particular site
        /// </summary>
        /// <param name="web">Web to process</param>
        /// <param name="rootWeb">The rootWeb</param>
        /// <param name="pageLayoutName">Name of the page layout</param>
        public static void SetDefaultPageLayoutForSite(this Web web, Web rootWeb, string pageLayoutName)
        {
            if (rootWeb == null)
                throw new ArgumentNullException(nameof(rootWeb));

            if (string.IsNullOrEmpty(pageLayoutName))
                throw new ArgumentNullException(nameof(pageLayoutName));

            // Save to property bag as the default page layout for the site
            XmlDocument xd = new XmlDocument();
            var node = CreateXmlNodeFromPageLayout(xd, web, rootWeb, pageLayoutName);
            web.SetPropertyBagValue(DefaultPageLayout, node.OuterXml);
        }

        private static XmlNode CreateXmlNodeFromPageLayout(XmlDocument xd, Web web, Web rootWeb, string pageLayoutName)
        {
            if (xd == null)
                throw new ArgumentNullException(nameof(xd));

            if (web == null)
                throw new ArgumentNullException(nameof(web));

            if (rootWeb == null)
                throw new ArgumentNullException(nameof(rootWeb));

            if (string.IsNullOrEmpty(pageLayoutName))
                throw new ArgumentNullException(nameof(pageLayoutName));

            var pageLayout = rootWeb.GetPageLayoutListItemByName(pageLayoutName);

            if (pageLayout != null)
            {
                // Parse the right styled xml for the layout - <layout guid="944ea6be-f287-42c6-aa11-3fd75ab1ee9e" url="_catalogs/masterpage/ArticleLeft.aspx" />
                XmlNode xmlNode = xd.CreateElement("layout");
                var xmlAttribute = xd.CreateAttribute("guid");
                xmlAttribute.Value = pageLayout["UniqueId"].ToString();
                var xmlAttribute2 = xd.CreateAttribute("url");
                // Get relative URL to the particular site collection
                xmlAttribute2.Value = SolveSiteRelativeUrl(rootWeb, pageLayout["FileRef"].ToString());
                xmlNode.Attributes.SetNamedItem(xmlAttribute);
                xmlNode.Attributes.SetNamedItem(xmlAttribute2);
                return xmlNode;
            }
            else
            {
                throw new IndexOutOfRangeException(String.Format(CoreResources.BrandingExtension_InvalidPageLayoutName, pageLayoutName));
            }
        }

        private static string SolveSiteRelativeUrl(Web web, string url)
        {
            if (web == null)
                throw new ArgumentNullException(nameof(web));

            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            web.EnsureProperties(w => w.ServerRelativeUrl);

            string newUrl = url.Substring(web.ServerRelativeUrl.Length);
            if (newUrl.Length > 0 && newUrl[0] == '/')
            {
                newUrl = newUrl.Substring(1);
            }
            return newUrl;
        }

        /// <summary>
        /// Can be used to set the site to inherit the default page layout option from parent. Cannot be used for root site of the site collection
        /// </summary>
        /// <param name="web">Web to operate against</param>
        public static void SetSiteToInheritPageLayouts(this Web web)
        {
            web.SetPropertyBagValue(DefaultPageLayout, Inherit);
        }

        /// <summary>
        /// Allow the web to use all available page layouts
        /// </summary>
        /// <param name="web">Web to operate against</param>
        public static void AllowAllPageLayouts(this Web web)
        {
            web.SetPropertyBagValue(AvailablePageLayouts, string.Empty);
        }

        /// <summary>
        /// Sets the available page layouts
        /// </summary>
        /// <param name="web">The web to process</param>
        /// <param name="rootWeb">The rootweb</param>
        /// <param name="pageLayouts">The page layouts to make available</param>
        public static void SetAvailablePageLayouts(this Web web, Web rootWeb, IEnumerable<string> pageLayouts)
        {
            var xd = new XmlDocument();
            XmlNode xmlNode = xd.CreateElement("pagelayouts");
            xd.AppendChild(xmlNode);
            foreach (var item in pageLayouts)
            {
                var node = CreateXmlNodeFromPageLayout(xd, web, rootWeb, item);
                xmlNode.AppendChild(node);
            }
            web.SetPropertyBagValue(AvailablePageLayouts, xmlNode.OuterXml);
        }

        /// <summary>
        /// Defines which templates are available for subsite creation
        /// </summary>
        /// <param name="web">Web to operate against</param>
        /// <param name="availableTemplates">List of <see cref="WebTemplateEntity"/> objects that define the templates that are allowed</param>
        public static void SetAvailableWebTemplates(this Web web, List<WebTemplateEntity> availableTemplates)
        {
            var propertyValue = string.Empty;

            var languages = new LanguageTemplateHash();
            foreach (var item in availableTemplates)
            {
                AddTemplateToCollection(languages, item);
            }

            if (availableTemplates.Count > 0)
            {
                var xd = new XmlDocument();
                XmlNode xmlNode = xd.CreateElement("webtemplates");
                xd.AppendChild(xmlNode);
                foreach (var language in languages)
                {
                    var xmlLcidNode = xmlNode.AppendChild(xd.CreateElement("lcid"));
                    var xmlAttribute = xd.CreateAttribute("id");
                    xmlAttribute.Value = language.Key;
                    xmlLcidNode.Attributes.SetNamedItem(xmlAttribute);

                    foreach (string item in language.Value)
                    {
                        var xmlWTNode = xmlLcidNode.AppendChild(xd.CreateElement("webtemplate"));
                        var xmlAttributeName = xd.CreateAttribute("name");
                        xmlAttributeName.Value = item;
                        xmlWTNode.Attributes.SetNamedItem(xmlAttributeName);
                    }
                }
                propertyValue = xmlNode.OuterXml;
            }
            //Save the xml entry to property bag
            web.SetPropertyBagValue(AvailableWebTemplates, propertyValue);
            //Set that templates are not inherited
            web.SetPropertyBagValue(InheritWebTemplates, "False");
        }

        /// <summary>
        /// Can be used to remote filters from the available web template
        /// </summary>
        /// <param name="web">Web to process</param>
        public static void ClearAvailableWebTemplates(this Web web)
        {
            web.SetPropertyBagValue(AvailableWebTemplates, "");
        }

        private static void AddTemplateToCollection(LanguageTemplateHash languages, WebTemplateEntity item)
        {
            var key = string.Empty;
            if (string.IsNullOrEmpty(item.LanguageCode))
            {
                key = "all";
            }
            else
            {
                key = item.LanguageCode;
            }

            if (!languages.ContainsKey(key))
            {
                languages[key] = new List<string>();
            }
            languages[key].Add(item.TemplateName);
        }

        /// <summary>
        /// Sets the web home page
        /// </summary>
        /// <param name="web">The Web to process</param>
        /// <param name="rootFolderRelativePath">The path relative to the root folder of the site, e.g. SitePages/Home.aspx</param>
        public static void SetHomePage(this Web web, string rootFolderRelativePath)
        {
            var folder = web.RootFolder;
            folder.WelcomePage = rootFolderRelativePath;
            folder.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Enables the responsive UI of a classic SharePoint Web
        /// </summary>
        /// <param name="web">The Web to activate the Responsive UI to</param>
        /// <param name="infrastructureUrl">URL pointing to an infrastructure site</param>
        [Obsolete]
        public static void EnableResponsiveUI(this Web web, string infrastructureUrl = null)
        {
            EnableResponsiveUIImplementation(web, infrastructureUrl);
        }

        /// <summary>
        /// Enables the responsive UI of a classic SharePoint Site
        /// </summary>
        /// <param name="site">The Site to activate the Responsive UI to</param>
        /// <param name="infrastructureUrl">URL pointing to an infrastructure site</param>
        [Obsolete]
        public static void EnableResponsiveUI(this Site site, string infrastructureUrl = null)
        {
            EnableResponsiveUIImplementation(site, infrastructureUrl);
        }

        /// <summary>
        /// Enables the responsive UI of a classic SharePoint Web or Site
        /// </summary>
        /// <param name="clientObject">The Web or Site to activate the Responsive UI to</param>
        /// <param name="infrastructureUrl">URL pointing to an infrastructure site</param>
        private static void EnableResponsiveUIImplementation(ClientObject clientObject, string infrastructureUrl = null)
        {
            // Double-check that we are targeting a Web or a Site
            if (clientObject is Web || clientObject is Site)
            {
                Web web = null;
                Site site = null;

                // If the target is a Web
                if (clientObject is Web)
                {
                    // Get it
                    web = ((Web)clientObject);
                }
                else
                {
                    // Otherwise get both the Site and the Web
                    site = ((Site)clientObject);
                    web = site.EnsureProperty(s => s.RootWeb);
                }

                if (web != null)
                {
                    web.EnsureProperty(w => w.ServerRelativeUrl);

                    var linkUrl = string.Empty;

                    if (!string.IsNullOrEmpty(infrastructureUrl))
                    {
                        using (var infrastructureContext = web.Context.Clone(infrastructureUrl))
                        {
                            var targetFolder = infrastructureContext.Web.EnsureFolderPath("Style Library/SP.Responsive.UI");
                            // Check if the file is there, if so, don't upload it.
                            var jsFile = targetFolder.GetFile("SP-Responsive-UI.js");
                            if (jsFile == null)
                            {
                                linkUrl = UploadStringAsFile(infrastructureContext.Web, targetFolder,
#if DEBUG || DEBUG15 || DEBUG16 || DEBUG19
                                    CoreResources.SP_Responsive_UI
#else
                                    CoreResources.SP_Responsive_UI_min
#endif
                                    , 
                                    "SP-Responsive-UI.js");
                            }
                            else
                            {
                                jsFile.EnsureProperty(f => f.ServerRelativeUrl);
                                linkUrl = jsFile.ServerRelativeUrl;
                            }

                            // Check if the file is there, if so, don't upload it.
                            if (targetFolder.GetFile("SP-Responsive-UI.css") == null)
                            {
                                UploadStringAsFile(infrastructureContext.Web, targetFolder,
#if DEBUG || DEBUG15 || DEBUG16 || DEBUG19
                                    CoreResources.SP_Responsive_UI_CSS
#else
                                    CoreResources.SP_Responsive_UI_CSS_min
#endif
                                    , 
                                    "SP-Responsive-UI.css");
                            }
                        }
                    }
                    else
                    {
                        var targetFolder = web.EnsureFolderPath("Style Library/SP.Responsive.UI");

                        linkUrl = UploadStringAsFile(web, targetFolder,
#if DEBUG || DEBUG15 || DEBUG16 || DEBUG19
                            CoreResources.SP_Responsive_UI
#else
                            CoreResources.SP_Responsive_UI_min
#endif
                            ,
                            "SP-Responsive-UI.js");
                        UploadStringAsFile(web, targetFolder,
#if DEBUG || DEBUG15 || DEBUG16 || DEBUG19
                            CoreResources.SP_Responsive_UI_CSS
#else
                            CoreResources.SP_Responsive_UI_CSS_min
#endif
                            , 
                            "SP-Responsive-UI.css");
                    }

                    // Deactive mobile feature
                    web.DeactivateFeature(new Guid("d95c97f3-e528-4da2-ae9f-32b3535fbb59"));
                    if (!string.IsNullOrEmpty(linkUrl))
                    {
                        if (site != null)
                        {
                            // If we have the Site enable the responsive UI site-wide
                            site.AddJsLink("PnPResponsiveUI", linkUrl, 0);
                        }
                        else
                        {
                            // Otherwise just target the Web
                            web.AddJsLink("PnPResponsiveUI", linkUrl, 0);
                        }
                    }
                }
            }
        }

        private static string UploadStringAsFile(Web web, Folder folder, string contents, string fileName)
        {
            var url = string.Empty;
            var targetFile = folder.GetFile(fileName);
            var checkedOut = false;
            if (targetFile != null)
            {
                CheckOutIfNeeded(web, targetFile);
            }
            using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(contents)))
            {
                var file = folder.UploadFile(fileName, stream, true);
                checkedOut = CheckOutIfNeeded(web, file);
                if (checkedOut)
                {
                    file.CheckIn("", CheckinType.MajorCheckIn);
                    web.Context.ExecuteQueryRetry();
                }
                file.EnsureProperty(f => f.ServerRelativeUrl);
                url = file.ServerRelativeUrl;
            }
            return url;
        }

        /// <summary>
        /// Disables the Responsive UI on a Classic SharePoint Web
        /// </summary>
        /// <param name="web">The Web to disable the Responsive UI on</param>
        [Obsolete]
        public static void DisableResponsiveUI(this Web web)
        {
            try
            {
                web.DeleteJsLink("PnPResponsiveUI");
            }
            catch
            {
                // Swallow exception as responsive UI might not be active.
            }
        }

        /// <summary>
        /// Disables the Responsive UI on a Classic SharePoint Site
        /// </summary>
        /// <param name="site">The Site to disable the Responsive UI on</param>
        [Obsolete]
        public static void DisableResponsiveUI(this Site site)
        {
            try
            {
                site.DeleteJsLink("PnPResponsiveUI");
            }
            catch
            {
                // Swallow exception as responsive UI might not be active.
            }
        }

        private static bool CheckOutIfNeeded(Web web, File targetFile)
        {
            var checkedOut = false;
            try
            {
                web.Context.Load(targetFile, f => f.CheckOutType, f => f.ListItemAllFields.ParentList.ForceCheckout);
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
#if !ONPREMISES
                if (ex.ServerErrorCode != -2113929210)
#else
                if (ex.ServerErrorCode != -2146232832)
#endif
                {
                    throw;
                }
            }
            return checkedOut;
        }


    }
}
