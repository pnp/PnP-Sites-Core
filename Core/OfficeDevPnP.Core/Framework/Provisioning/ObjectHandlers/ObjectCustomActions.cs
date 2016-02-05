using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Diagnostics;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectCustomActions : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Custom Actions"; }
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                var site = context.Site;

                // if this is a sub site then we're not enabling the site collection scoped custom actions
                if (!web.IsSubSite())
                {
                    var siteCustomActions = template.CustomActions.SiteCustomActions;
                    ProvisionCustomActionImplementation(site, siteCustomActions, parser, scope);
                }

                var webCustomActions = template.CustomActions.WebCustomActions;
                ProvisionCustomActionImplementation(web, webCustomActions, parser, scope);

                // Switch parser context back to it's original context
                parser.Rebase(web);
            }
            return parser;
        }

        private void ProvisionCustomActionImplementation(object parent, CustomActionCollection customActions, TokenParser parser, PnPMonitoredScope scope)
        {
            Web web = null;
            Site site = null;
            if (parent is Site)
            {
                site = parent as Site;

                // Switch parser context;
                parser.Rebase(site.RootWeb);
            }
            else
            {
                web = parent as Web;

                // Switch parser context
                parser.Rebase(web);
            }
            foreach (var customAction in customActions)
            {
                var caExists = false;
                if (site != null)
                {
                    caExists = site.CustomActionExists(customAction.Name);
                }
                else
                {
                    caExists = web.CustomActionExists(customAction.Name);
                }
                if (!caExists)
                {
                    var customActionEntity = new CustomActionEntity()
                    {
                        CommandUIExtension = customAction.CommandUIExtension != null ? parser.ParseString(customAction.CommandUIExtension.ToString()) : string.Empty,
                        Description = customAction.Description,
                        Group = customAction.Group,
                        ImageUrl = parser.ParseString(customAction.ImageUrl),
                        Location = customAction.Location,
                        Name = customAction.Name,
                        RegistrationId = customAction.RegistrationId,
                        RegistrationType = customAction.RegistrationType,
                        Remove = customAction.Remove,
                        Rights = customAction.Rights,
                        ScriptBlock = parser.ParseString(customAction.ScriptBlock),
                        ScriptSrc = parser.ParseString(customAction.ScriptSrc, "~site", "~sitecollection"),
                        Sequence = customAction.Sequence,
                        Title = customAction.Title,
                        Url = parser.ParseString(customAction.Url)
                    };


                    if (site != null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_custom_action___0___to_scope_Site, customActionEntity.Name);
                        site.AddCustomAction(customActionEntity);
#if !CLIENTSDKV15
                        if (customAction.Title.ContainsResourceToken() || customAction.Description.ContainsResourceToken())
                        {
                            var uca = site.GetCustomActions().Where(uc => uc.Name == customAction.Name).FirstOrDefault();
                            SetCustomActionResourceValues(parser, customAction, uca);
                        }
#endif
                    }
                    else
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_custom_action___0___to_scope_Web, customActionEntity.Name);
                        web.AddCustomAction(customActionEntity);
#if !CLIENTSDKV15
                        if (customAction.Title.ContainsResourceToken() || customAction.Description.ContainsResourceToken())
                        {
                            var uca = web.GetCustomActions().Where(uc => uc.Name == customAction.Name).FirstOrDefault();
                            SetCustomActionResourceValues(parser, customAction, uca);
                        }
#endif
                    }
                }
                else
                {
                    UserCustomAction existingCustomAction = null;
                    if (site != null)
                    {
                        existingCustomAction = site.GetCustomActions().FirstOrDefault(c => c.Name == customAction.Name);
                    }
                    else
                    {
                        existingCustomAction = web.GetCustomActions().FirstOrDefault(c => c.Name == customAction.Name);
                    }
                    if (existingCustomAction != null)
                    {
                        var isDirty = false;

                        if (customAction.CommandUIExtension != null)
                        {
                            if (existingCustomAction.CommandUIExtension != parser.ParseString(customAction.CommandUIExtension.ToString()))
                            {
                                scope.LogPropertyUpdate("CommandUIExtension");
                                existingCustomAction.CommandUIExtension = parser.ParseString(customAction.CommandUIExtension.ToString());
                                isDirty = true;
                            }
                        }

                        if (existingCustomAction.Description != customAction.Description)
                        {
                            scope.LogPropertyUpdate("Description");
                            existingCustomAction.Description = customAction.Description;
                            isDirty = true;
                        }
#if !CLIENTSDKV15
                        if (customAction.Description.ContainsResourceToken())
                        {
                            if (existingCustomAction.DescriptionResource.SetUserResourceValue(customAction.Description, parser))
                            {
                                isDirty = true;
                            }
                        }
#endif
                        if (existingCustomAction.Group != customAction.Group)
                        {
                            scope.LogPropertyUpdate("Group");
                            existingCustomAction.Group = customAction.Group;
                            isDirty = true;
                        }
                        if (existingCustomAction.ImageUrl != parser.ParseString(customAction.ImageUrl))
                        {
                            scope.LogPropertyUpdate("ImageUrl");
                            existingCustomAction.ImageUrl = parser.ParseString(customAction.ImageUrl);
                            isDirty = true;
                        }
                        if (existingCustomAction.Location != customAction.Location)
                        {
                            scope.LogPropertyUpdate("Location");
                            existingCustomAction.Location = customAction.Location;
                            isDirty = true;
                        }
                        if (existingCustomAction.RegistrationId != customAction.RegistrationId)
                        {
                            scope.LogPropertyUpdate("RegistrationId");
                            existingCustomAction.RegistrationId = customAction.RegistrationId;
                            isDirty = true;
                        }
                        if (existingCustomAction.RegistrationType != customAction.RegistrationType)
                        {
                            scope.LogPropertyUpdate("RegistrationType");
                            existingCustomAction.RegistrationType = customAction.RegistrationType;
                            isDirty = true;
                        }
                        if (existingCustomAction.ScriptBlock != parser.ParseString(customAction.ScriptBlock))
                        {
                            scope.LogPropertyUpdate("ScriptBlock");
                            existingCustomAction.ScriptBlock = parser.ParseString(customAction.ScriptBlock);
                            isDirty = true;
                        }
                        if (existingCustomAction.ScriptSrc != parser.ParseString(customAction.ScriptSrc, "~site", "~sitecollection"))
                        {
                            scope.LogPropertyUpdate("ScriptSrc");
                            existingCustomAction.ScriptSrc = parser.ParseString(customAction.ScriptSrc, "~site", "~sitecollection");
                            isDirty = true;
                        }
                        if (existingCustomAction.Sequence != customAction.Sequence)
                        {
                            scope.LogPropertyUpdate("Sequence");
                            existingCustomAction.Sequence = customAction.Sequence;
                            isDirty = true;
                        }
                        if (existingCustomAction.Title != parser.ParseString(customAction.Title))
                        {
                            scope.LogPropertyUpdate("Title");
                            existingCustomAction.Title = parser.ParseString(customAction.Title);
                            isDirty = true;
                        }
#if !CLIENTSDKV15
                        if (customAction.Title.ContainsResourceToken())
                        {
                            if (existingCustomAction.TitleResource.SetUserResourceValue(customAction.Title, parser))
                            {
                                isDirty = true;
                            }

                        }
#endif
                        if (existingCustomAction.Url != parser.ParseString(customAction.Url))
                        {
                            scope.LogPropertyUpdate("Url");
                            existingCustomAction.Url = parser.ParseString(customAction.Url);
                            isDirty = true;
                        }
                        if (isDirty)
                        {
                            existingCustomAction.Update();
                            existingCustomAction.Context.ExecuteQueryRetry();
                        }
                    }
                }
            }
        }

        private static void SetCustomActionResourceValues(TokenParser parser, CustomAction customAction, UserCustomAction uca)
        {
            if (uca != null)
            {
                bool isDirty = false;
#if !CLIENTSDKV15
                if (customAction.Title.ContainsResourceToken())
                {
                    if (uca.TitleResource.SetUserResourceValue(customAction.Title, parser))
                    {
                        isDirty = true;
                    }
                }
                if (customAction.Description.ContainsResourceToken())
                {
                    if (uca.DescriptionResource.SetUserResourceValue(customAction.Description, parser))
                    {
                        isDirty = true;
                    }
                }
#endif
                if (isDirty)
                {
                    uca.Update();
                    uca.Context.ExecuteQueryRetry();
                }
            }
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                bool isSubSite = web.IsSubSite();
                var webCustomActions = web.GetCustomActions();
                var siteCustomActions = context.Site.GetCustomActions();

                var customActions = new CustomActions();
                foreach (var customAction in webCustomActions)
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_web_scoped_custom_action___0___to_template, customAction.Name);
                    customActions.WebCustomActions.Add(CopyUserCustomAction(customAction));
                }

                // if this is a sub site then we're not creating entities for site collection scoped custom actions
                if (!isSubSite)
                {
                    foreach (var customAction in siteCustomActions)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_site_scoped_custom_action___0___to_template, customAction.Name);
                        customActions.SiteCustomActions.Add(CopyUserCustomAction(customAction));
                    }
                }

                template.CustomActions = customActions;

                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate, isSubSite, scope);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate, bool isSubSite, PnPMonitoredScope scope)
        {
            if (!isSubSite)
            {
                foreach (var customAction in baseTemplate.CustomActions.SiteCustomActions)
                {
                    int index = template.CustomActions.SiteCustomActions.FindIndex(f => f.Name.Equals(customAction.Name));

                    if (index > -1)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Removing_site_scoped_custom_action___0___from_template_because_already_available_in_base_template, customAction.Name);
                        template.CustomActions.SiteCustomActions.RemoveAt(index);
                    }
                }
            }

            foreach (var customAction in baseTemplate.CustomActions.WebCustomActions)
            {
                int index = template.CustomActions.WebCustomActions.FindIndex(f => f.Name.Equals(customAction.Name));

                if (index > -1)
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Removing_web_scoped_custom_action___0___from_template_because_already_available_in_base_template, customAction.Name);
                    template.CustomActions.WebCustomActions.RemoveAt(index);
                }
            }

            return template;
        }

        private CustomAction CopyUserCustomAction(UserCustomAction userCustomAction)
        {
            var customAction = new CustomAction();
            customAction.Description = userCustomAction.Description;
            customAction.Enabled = true;
            customAction.Group = userCustomAction.Group;
            customAction.ImageUrl = userCustomAction.ImageUrl;
            customAction.Location = userCustomAction.Location;
            customAction.Name = userCustomAction.Name;
            customAction.Rights = userCustomAction.Rights;
            customAction.ScriptBlock = userCustomAction.ScriptBlock;
            customAction.ScriptSrc = userCustomAction.ScriptSrc;
            customAction.Sequence = userCustomAction.Sequence;
            customAction.Title = userCustomAction.Title;
            customAction.Url = userCustomAction.Url;
            customAction.RegistrationId = userCustomAction.RegistrationId;
            customAction.RegistrationType = userCustomAction.RegistrationType;
            customAction.CommandUIExtension = !System.String.IsNullOrEmpty(userCustomAction.CommandUIExtension) ?
                XElement.Parse(userCustomAction.CommandUIExtension) : null;

            return customAction;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.CustomActions != null ? template.CustomActions.SiteCustomActions.Any() || template.CustomActions.WebCustomActions.Any() : false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var context = web.Context as ClientContext;
                var webCustomActions = web.GetCustomActions();
                var siteCustomActions = context.Site.GetCustomActions();

                _willExtract = webCustomActions.Any() || siteCustomActions.Any();
            }
            return _willExtract.Value;
        }
    }
}
