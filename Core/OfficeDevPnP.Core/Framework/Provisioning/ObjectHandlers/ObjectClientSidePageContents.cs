using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Utilities;
using System;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
#if !ONPREMISES
    internal class ObjectClientSidePageContents: ObjectContentHandlerBase
    {
        public override string Name
        {
            get { return "Client Side Page Contents"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // This handler only extracts contents and adds them to the Files and Pages collection.
            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Extract the Home Page
                web.EnsureProperties(w => w.RootFolder.WelcomePage, w => w.ServerRelativeUrl, w => w.Url);

                var homePageUrl = web.RootFolder.WelcomePage;
                var homepageName = System.IO.Path.GetFileName(web.RootFolder.WelcomePage);
                if (string.IsNullOrEmpty(homepageName))
                {
                    homepageName = "Home.aspx";
                }

                try
                {
                    var homePage = web.LoadClientSidePage(homepageName);

                    if (homePage.Sections.Count == 0 && homePage.Controls.Count == 0)
                    {
                        // This is default home page which was not customized...and as such there's no page definition stored in the list item. We don't need to extact this page.
                        scope.LogInfo(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_DefaultHomePage);
                    }
                    else
                    {
                        // Create the page
                        var homePageInstance = new ClientSidePage()
                        {
                            PageName = homepageName,
                            PromoteAsNewsArticle = false,
                            Overwrite = false,
                        };

                        // Add the sections
                        foreach (var section in homePage.Sections)
                        {
                            // Set order
                            var sectionInstance = new CanvasSection()
                            {
                                Order = section.Order,
                            };

                            // Set section type
                            switch (section.Type)
                            {
                                case Pages.CanvasSectionTemplate.OneColumn:
                                    sectionInstance.Type = CanvasSectionType.OneColumn;
                                    break;
                                case Pages.CanvasSectionTemplate.TwoColumn:
                                    sectionInstance.Type = CanvasSectionType.TwoColumn;
                                    break;
                                case Pages.CanvasSectionTemplate.TwoColumnLeft:
                                    sectionInstance.Type = CanvasSectionType.TwoColumnLeft;
                                    break;
                                case Pages.CanvasSectionTemplate.TwoColumnRight:
                                    sectionInstance.Type = CanvasSectionType.TwoColumnRight;
                                    break;
                                case Pages.CanvasSectionTemplate.ThreeColumn:
                                    sectionInstance.Type = CanvasSectionType.ThreeColumn;
                                    break;
                                case Pages.CanvasSectionTemplate.OneColumnFullWidth:
                                    sectionInstance.Type = CanvasSectionType.OneColumnFullWidth;
                                    break;
                                default:
                                    sectionInstance.Type = CanvasSectionType.OneColumn;
                                    break;
                            }

                            // Add controls to section
                            foreach (var column in section.Columns)
                            {
                                foreach (var control in column.Controls)
                                {
                                    // Create control 
                                    CanvasControl controlInstance = new CanvasControl()
                                    {
                                        Column = column.Order,
                                        ControlId = control.InstanceId,
                                        Order = control.Order,
                                    };

                                    // Set control type
                                    if (control.Type == typeof(Pages.ClientSideText))
                                    {
                                        controlInstance.Type = WebPartType.Text;

                                        // Set text content
                                        controlInstance.ControlProperties = new System.Collections.Generic.Dictionary<string, string>(1)
                                        {
                                            { "Text", (control as Pages.ClientSideText).Text }
                                        };
                                    }
                                    else
                                    {
                                        // set ControlId to webpart id 
                                        controlInstance.ControlId = Guid.Parse((control as Pages.ClientSideWebPart).WebPartId);
                                        var webPartType = Pages.ClientSidePage.NameToClientSideWebPartEnum((control as Pages.ClientSideWebPart).WebPartId);
                                        switch (webPartType)
                                        {
                                            case Pages.DefaultClientSideWebParts.ContentRollup:
                                                controlInstance.Type = WebPartType.ContentRollup;
                                                break;
                                            case Pages.DefaultClientSideWebParts.BingMap:
                                                controlInstance.Type = WebPartType.BingMap;
                                                break;
                                            case Pages.DefaultClientSideWebParts.ContentEmbed:
                                                controlInstance.Type = WebPartType.ContentEmbed;
                                                break;
                                            case Pages.DefaultClientSideWebParts.DocumentEmbed:
                                                controlInstance.Type = WebPartType.DocumentEmbed;
                                                break;
                                            case Pages.DefaultClientSideWebParts.Image:
                                                controlInstance.Type = WebPartType.Image;
                                                break;
                                            case Pages.DefaultClientSideWebParts.ImageGallery:
                                                controlInstance.Type = WebPartType.ImageGallery;
                                                break;
                                            case Pages.DefaultClientSideWebParts.LinkPreview:
                                                controlInstance.Type = WebPartType.LinkPreview;
                                                break;
                                            case Pages.DefaultClientSideWebParts.NewsFeed:
                                                controlInstance.Type = WebPartType.NewsFeed;
                                                break;
                                            case Pages.DefaultClientSideWebParts.NewsReel:
                                                controlInstance.Type = WebPartType.NewsReel;
                                                break;
                                            case Pages.DefaultClientSideWebParts.PowerBIReportEmbed:
                                                controlInstance.Type = WebPartType.PowerBIReportEmbed;
                                                break;
                                            case Pages.DefaultClientSideWebParts.QuickChart:
                                                controlInstance.Type = WebPartType.QuickChart;
                                                break;
                                            case Pages.DefaultClientSideWebParts.SiteActivity:
                                                controlInstance.Type = WebPartType.SiteActivity;
                                                break;
                                            case Pages.DefaultClientSideWebParts.VideoEmbed:
                                                controlInstance.Type = WebPartType.VideoEmbed;
                                                break;
                                            case Pages.DefaultClientSideWebParts.YammerEmbed:
                                                controlInstance.Type = WebPartType.YammerEmbed;
                                                break;
                                            case Pages.DefaultClientSideWebParts.Events:
                                                controlInstance.Type = WebPartType.Events;
                                                break;
                                            case Pages.DefaultClientSideWebParts.GroupCalendar:
                                                controlInstance.Type = WebPartType.GroupCalendar;
                                                break;
                                            case Pages.DefaultClientSideWebParts.Hero:
                                                controlInstance.Type = WebPartType.Hero;
                                                break;
                                            case Pages.DefaultClientSideWebParts.List:
                                                controlInstance.Type = WebPartType.List;
                                                break;
                                            case Pages.DefaultClientSideWebParts.PageTitle:
                                                controlInstance.Type = WebPartType.PageTitle;
                                                break;
                                            case Pages.DefaultClientSideWebParts.People:
                                                controlInstance.Type = WebPartType.People;
                                                break;
                                            case Pages.DefaultClientSideWebParts.QuickLinks:
                                                controlInstance.Type = WebPartType.QuickLinks;
                                                break;
                                            case Pages.DefaultClientSideWebParts.ThirdParty:
                                                controlInstance.Type = WebPartType.Custom;
                                                break;
                                            default:
                                                controlInstance.Type = WebPartType.Custom;
                                                break;
                                        }

                                        // set the control properties
                                        controlInstance.JsonControlData = (control as Pages.ClientSideWebPart).PropertiesJson;
                                    }

                                    // add control to section
                                    sectionInstance.Controls.Add(controlInstance);
                                }
                            }

                            homePageInstance.Sections.Add(sectionInstance);
                        }

                        // Renumber the sections...when editing modern homepages you can end up with section with order 0.5 or 0.75...let's ensure we render section as of 1
                        int sectionOrder = 1;
                        foreach(var sectionInstance in homePageInstance.Sections)
                        {
                            sectionInstance.Order = sectionOrder;
                            sectionOrder++;
                        }

                        // Add the page to the template
                        template.ClientSidePages.Add(homePageInstance);

                        // Set the homepage
                        if (template.WebSettings == null)
                        {
                            template.WebSettings = new WebSettings();
                        }
                        template.WebSettings.WelcomePage = homePageUrl;
                    }
                }
                catch (ArgumentException ex)
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_ClientSidePageContents_NoValidPage, ex.Message);
                }

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
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
