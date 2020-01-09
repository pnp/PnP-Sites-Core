using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
#if !SP2013 && !SP2016
    internal class ObjectClientSidePageContents : ObjectContentHandlerBase
    {
        
        public override string Name
        {
            get { return "Client Side Page Contents"; }
        }

        public override string InternalName => "ClientSidePageContents";
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // This handler only extracts contents and adds them to the Files and Pages collection.
            return parser;
        }

        private const string CAMLQueryByExtension = @"
                <View Scope='Recursive'>
                  <Query>
                    <Where>
                      <Contains>
                        <FieldRef Name='File_x0020_Type'/>
                        <Value Type='text'>aspx</Value>
                      </Contains>
                    </Where>
                  </Query>
                </View>";

        private const string FileRefField = "FileRef";
        private const string FileLeafRefField = "FileLeafRef";
        private const string ClientSideApplicationId = "ClientSideApplicationId";
        private const string SPIsTranslation = "_SPIsTranslation";
        private const string SPTranslatedLanguages = "_SPTranslatedLanguages";
        private const string PageIDField = "UniqueId";
        private const string SPTranslationSourceItemId = "_SPTranslationSourceItemId";
        private const string SPTranslationLanguage = "_SPTranslationLanguage";

        private static readonly Guid FeatureId_Web_ModernPage = new Guid("B6917CB1-93A0-4B97-A84D-7CF49975D4EC");
        public const string TemplatesFolderGuid = "vti_TemplatesFolderGuid";

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var clientSidePageContentsHelper = new ClientSidePageContentsHelper();

                var baseUrl = web.EnsureProperty(w => w.ServerRelativeUrl) + "/SitePages/";

                // Extract the Home Page
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);
                var homePageUrl = web.RootFolder.WelcomePage;

                // Get pages library
                List sitePagesLibrary = null;
                try
                {
                    ListCollection listCollection = web.Lists;
                    listCollection.EnsureProperties(coll => coll.Include(li => li.BaseTemplate, li => li.RootFolder));
                    sitePagesLibrary = listCollection.Where(p => p.BaseTemplate == (int)ListTemplateType.WebPageLibrary).FirstOrDefault();
                } 
                catch
                {
                    // fall back in case of exception when the site has been incorrectly provisioned which can cause access issues on lists/libraries.
                    sitePagesLibrary = web.Lists.GetByTitle("Site Pages");
                    sitePagesLibrary.EnsureProperties(l => l.BaseTemplate, l => l.RootFolder);
                }

                if (sitePagesLibrary != null)
                {
                    var templateFolderName = OfficeDevPnP.Core.Pages.ClientSidePage.DefaultTemplatesFolder;// string.Empty;
                    var templateFolderString = sitePagesLibrary.GetPropertyBagValueString(TemplatesFolderGuid, null);
                    Guid.TryParse(templateFolderString, out Guid templateFolderGuid);
                    if (templateFolderGuid != Guid.Empty)
                    {
                        try
                        {
                            var templateFolder = ((ClientContext)sitePagesLibrary.Context).Web.GetFolderById(templateFolderGuid);
                            templateFolderName = templateFolder.EnsureProperty(f => f.Name);
                        }
                        catch
                        {
                            //eat it and continue with default name
                        }
                    }
                    CamlQuery query = new CamlQuery
                    {
                        ViewXml = CAMLQueryByExtension
                    };
                    var pages = sitePagesLibrary.GetItems(query);
                    web.Context.Load(pages);
                    web.Context.ExecuteQueryRetry();
                    if (pages.FirstOrDefault() != null)
                    {
                        // Prep a list of pages to export allowing us hanlde translations
                        List<PageToExport> pagesToExport = new List<PageToExport>();
                        foreach(var page in pages)
                        {
                            PageToExport pageToExport = new PageToExport()
                            {
                                ListItem = page,
                                IsTranslation = false,
                                TranslatedLanguages = null,
                            };

                            // If multi-lingual is enabled these fields will be available on the SitePages library
                            if (page.FieldValues.ContainsKey(SPIsTranslation) && page[SPIsTranslation] != null && !string.IsNullOrEmpty(page[SPIsTranslation].ToString()))
                            {
                                if (bool.TryParse(page[SPIsTranslation].ToString(), out bool isTranslation))
                                {
                                    pageToExport.IsTranslation = isTranslation;
                                }
                            }

                            if (page.FieldValues.ContainsKey(PageIDField) && page[PageIDField] != null && !string.IsNullOrEmpty(page[PageIDField].ToString()))
                            {
                                pageToExport.PageId = Guid.Parse(page[PageIDField].ToString());
                            }

                            if (page.FieldValues.ContainsKey(SPTranslationSourceItemId) && page[SPTranslationSourceItemId] != null && !string.IsNullOrEmpty(page[SPTranslationSourceItemId].ToString()))
                            {
                                pageToExport.SourcePageId = Guid.Parse(page[SPTranslationSourceItemId].ToString());
                            }

                            if (page.FieldValues.ContainsKey(SPTranslationLanguage) && page[SPTranslationLanguage] != null && !string.IsNullOrEmpty(page[SPTranslationLanguage].ToString()))
                            {
                                pageToExport.Language = page[SPTranslationLanguage].ToString();
                            }

                            if (page.FieldValues.ContainsKey(SPTranslatedLanguages) && page[SPTranslatedLanguages] != null && !string.IsNullOrEmpty(page[SPTranslatedLanguages].ToString()))
                            {
                                pageToExport.TranslatedLanguages = new List<string>(page[SPTranslatedLanguages] as string[]);
                            }

                            string pageUrl = null;
                            string pageName = "";
                            if (page.FieldValues.ContainsKey(FileRefField) && !String.IsNullOrEmpty(page[FileRefField].ToString()))
                            {
                                pageUrl = page[FileRefField].ToString();
                                pageName = page[FileLeafRefField].ToString();
                            }
                            else
                            {
                                //skip page
                                continue;
                            }

                            var isTemplate = false;
                            // Is this page a template?
                            if (pageUrl.IndexOf($"/{templateFolderName}/", StringComparison.InvariantCultureIgnoreCase) > -1)
                            {
                                isTemplate = true;
                            }
                            // Is this page the web's home page?
                            bool isHomePage = false;
                            if (pageUrl.EndsWith(homePageUrl, StringComparison.InvariantCultureIgnoreCase))
                            {
                                isHomePage = true;
                            }

                            // Get the name of the page, including the folder name
                            pageName = Regex.Replace(pageUrl, baseUrl, "", RegexOptions.IgnoreCase);

                            pageToExport.IsHomePage = isHomePage;
                            pageToExport.IsTemplate = isTemplate;
                            pageToExport.PageName = pageName;
                            pageToExport.PageUrl = pageUrl;
                            pagesToExport.Add(pageToExport);
                        }

                        // Populate SourcePageName to make it easier to hookup translations at export time
                        foreach (var page in pagesToExport.Where(p => p.IsTranslation))
                        {
                            var sourcePage = pagesToExport.Where(p => p.PageId == page.SourcePageId).FirstOrDefault();
                            if (sourcePage != null)
                            {
                                page.SourcePageName = sourcePage.PageName;
                            }
                        }

                        var currentPageIndex = 1;
                        foreach (var page in pagesToExport.OrderBy(p=>p.IsTranslation))
                        {
                            if (creationInfo.IncludeAllClientSidePages || page.IsHomePage)
                            {
                                // Is this a client side page?
                                if (FieldExistsAndUsed(page.ListItem, ClientSideApplicationId) && page.ListItem[ClientSideApplicationId].ToString().Equals(FeatureId_Web_ModernPage.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                {
                                    WriteSubProgress("ClientSidePage", !string.IsNullOrWhiteSpace(page.PageName) ? page.PageName : page.PageUrl, currentPageIndex, pages.Count);
                                    // extract the page using the OOB logic
                                    clientSidePageContentsHelper.ExtractClientSidePage(web, template, creationInfo, scope, page);
                                }
                            }
                            currentPageIndex++;
                        }
                    }
                }

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }

            }
            return template;
        }

        private static bool FieldExistsAndUsed(ListItem item, string fieldName)
        {
            return (item.FieldValues.ContainsKey(fieldName) && item[fieldName] != null);
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }

#endif
}
