using System;
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

        public override string InternalName => "CustomActions";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                var context = web.Context as ClientContext;
                var site = context.Site;

                // Check if this is not a noscript site as we're not allowed to update some properties
                bool isNoScriptSite = web.IsNoScriptSite();

                // if this is a sub site then we're not enabling the site collection scoped custom actions
                if (!web.IsSubSite())
                {
                    var siteCustomActions = template.CustomActions.SiteCustomActions;
                    ProvisionCustomActionImplementation(site, siteCustomActions, parser, scope, isNoScriptSite);
                }

                var webCustomActions = template.CustomActions.WebCustomActions;
                ProvisionCustomActionImplementation(web, webCustomActions, parser, scope, isNoScriptSite);

                // Switch parser context back to it's original context
                parser.Rebase(web);
            }
            return parser;
        }

        private void ProvisionCustomActionImplementation(object parent, CustomActionCollection customActions, TokenParser parser, PnPMonitoredScope scope, bool isNoScriptSite= false)
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

                if (isNoScriptSite && Guid.Empty == customAction.ClientSideComponentId)
                {
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_CustomActions_SkippingAddUpdateDueToNoScript, customAction.Name);
                    continue;
                }

                var caExists = false;
                if (site != null)
                {
                    caExists = site.CustomActionExists(customAction.Name);
                }
                else
                {
                    caExists = web.CustomActionExists(customAction.Name);
                }

                // If the CustomAction does not exist, we don't have to remove it, and it is enabled
                if (!caExists && !customAction.Remove && customAction.Enabled)
                {
                    // Then we add it to the target
                    var customActionEntity = new CustomActionEntity()
                    {
#if !SP2013 && !SP2016
                        ClientSideComponentId = customAction.ClientSideComponentId,
                        ClientSideComponentProperties = customAction.ClientSideComponentProperties != null ? parser.ParseString(customAction.ClientSideComponentProperties) : customAction.ClientSideComponentProperties,
#endif
                        CommandUIExtension = customAction.CommandUIExtension != null ? parser.ParseString(customAction.CommandUIExtension.ToString()) : string.Empty,
                        Description = parser.ParseString(customAction.Description),
                        Group = customAction.Group,
                        ImageUrl = parser.ParseString(customAction.ImageUrl),
                        Location = customAction.Location,
                        Name = customAction.Name,
                        RegistrationId = parser.ParseString(customAction.RegistrationId),
                        RegistrationType = customAction.RegistrationType,
                        Remove = customAction.Remove,
                        Rights = customAction.Rights,
                        ScriptBlock = parser.ParseString(customAction.ScriptBlock),
                        ScriptSrc = parser.ParseString(customAction.ScriptSrc),
                        Sequence = customAction.Sequence,
                        Title = parser.ParseString(customAction.Title),
                        Url = parser.ParseString(customAction.Url)
                    };


                    if (site != null)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_custom_action___0___to_scope_Site, customActionEntity.Name);
                        site.AddCustomAction(customActionEntity);
#if !SP2013 && !SP2016
                        if ((!string.IsNullOrEmpty(customAction.Title) && customAction.Title.ContainsResourceToken()) ||
                            (!string.IsNullOrEmpty(customAction.Description) && customAction.Description.ContainsResourceToken()))
                        {
                            var uca = site.GetCustomActions().FirstOrDefault(uc => uc.Name == customAction.Name);
                            SetCustomActionResourceValues(parser, customAction, uca);
                        }
#endif
                    }
                    else
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_custom_action___0___to_scope_Web, customActionEntity.Name);
                        web.AddCustomAction(customActionEntity);
#if !SP2013 && !SP2016
                        if (customAction.Title.ContainsResourceToken() || customAction.Description.ContainsResourceToken())
                        {
                            var uca = web.GetCustomActions().FirstOrDefault(uc => uc.Name == customAction.Name);
                            SetCustomActionResourceValues(parser, customAction, uca);
                        }
#endif
                    }
                }
                else
                {
                    UserCustomAction existingCustomAction;
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
                        // If we have to remove the existing CustomAction
                        if (customAction.Remove)
                        {
                            // We simply remove it
                            existingCustomAction.DeleteObject();
                            existingCustomAction.Context.ExecuteQueryRetry();
                        }
                        else
                        {
                            UpdateCustomAction(parser, scope, customAction, existingCustomAction, isNoScriptSite);
                        }
                    }
                }
            }
        }

        internal static void UpdateCustomAction(TokenParser parser, PnPMonitoredScope scope, CustomAction customAction, UserCustomAction existingCustomAction, bool isNoScriptSite = false)
        {
            var isDirty = false;

            if (isNoScriptSite)
            {
                scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_CustomActions_SkippingAddUpdateDueToNoScript, customAction.Name);
                return;
            }

            // Otherwise we update it
            if (customAction.CommandUIExtension != null)
            {
                if (existingCustomAction.CommandUIExtension != parser.ParseString(customAction.CommandUIExtension.ToString()))
                {
                    scope.LogPropertyUpdate("CommandUIExtension");
                    existingCustomAction.CommandUIExtension = parser.ParseString(customAction.CommandUIExtension.ToString());
                    isDirty = true;
                }
            }
            else
            {
                // Required to allow for a delta action to blank out the CommandUIExtension attribute
                if (existingCustomAction.CommandUIExtension != null)
                {
                    scope.LogPropertyUpdate("CommandUIExtension");
                    existingCustomAction.CommandUIExtension = null;
                    isDirty = true;
                }
            }

#if !SP2013 && !SP2016
            if (customAction.ClientSideComponentId != null && customAction.ClientSideComponentId != Guid.Empty)
            {
                if  (existingCustomAction.ClientSideComponentId != customAction.ClientSideComponentId)
                {
                    existingCustomAction.ClientSideComponentId = customAction.ClientSideComponentId;
                    isDirty = true;
                }
            }

            if (!String.IsNullOrEmpty(customAction.ClientSideComponentProperties))
            {
                if (existingCustomAction.ClientSideComponentProperties != parser.ParseString(customAction.ClientSideComponentProperties))
                {
                    existingCustomAction.ClientSideComponentProperties = parser.ParseString(customAction.ClientSideComponentProperties);
                    isDirty = true;
                }
            }
#endif

            if (existingCustomAction.Description != customAction.Description)
            {
                scope.LogPropertyUpdate("Description");
                existingCustomAction.Description = customAction.Description;
                isDirty = true;
            }
#if !SP2013 && !SP2016
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
            if (existingCustomAction.RegistrationId != parser.ParseString(customAction.RegistrationId))
            {
                scope.LogPropertyUpdate("RegistrationId");
                existingCustomAction.RegistrationId = parser.ParseString(customAction.RegistrationId);
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
            if (existingCustomAction.ScriptSrc != parser.ParseString(customAction.ScriptSrc))
            {
                scope.LogPropertyUpdate("ScriptSrc");
                existingCustomAction.ScriptSrc = parser.ParseString(customAction.ScriptSrc);
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
#if !SP2013 && !SP2016
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

        private static void SetCustomActionResourceValues(TokenParser parser, CustomAction customAction, UserCustomAction uca)
        {
            if (uca != null)
            {
                bool isDirty = false;
#if !SP2013 && !SP2016
                if (!string.IsNullOrEmpty(customAction.Title) && customAction.Title.ContainsResourceToken())
                {
                    if (uca.TitleResource.SetUserResourceValue(customAction.Title, parser))
                    {
                        isDirty = true;
                    }
                }
                if (!string.IsNullOrEmpty(customAction.Description) && customAction.Description.ContainsResourceToken())
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
                var context = (ClientContext)web.Context;
                bool isSubSite = web.IsSubSite();
                var webCustomActions = web.GetCustomActions();
                var siteCustomActions = context.Site.GetCustomActions();

                var customActions = new CustomActions();
                foreach (var customAction in webCustomActions)
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_web_scoped_custom_action___0___to_template, customAction.Name);
                    customActions.WebCustomActions.Add(CopyUserCustomAction(customAction, creationInfo,template));
                }

                // if this is a sub site then we're not creating entities for site collection scoped custom actions
                if (!isSubSite)
                {
                    foreach (var customAction in siteCustomActions)
                    {
                        scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_CustomActions_Adding_site_scoped_custom_action___0___to_template, customAction.Name);
                        customActions.SiteCustomActions.Add(CopyUserCustomAction(customAction, creationInfo,template));
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

        private CustomAction CopyUserCustomAction(UserCustomAction userCustomAction, ProvisioningTemplateCreationInformation creationInfo, ProvisioningTemplate template)
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

#if !SP2013 && !SP2016
            customAction.ClientSideComponentId = userCustomAction.ClientSideComponentId;
            customAction.ClientSideComponentProperties = userCustomAction.ClientSideComponentProperties;
#endif

            customAction.CommandUIExtension = !System.String.IsNullOrEmpty(userCustomAction.CommandUIExtension) ?
                XElement.Parse(userCustomAction.CommandUIExtension) : null;

#if !SP2013 && !SP2016
            if (creationInfo.PersistMultiLanguageResources)
            {
                var resourceKey = userCustomAction.Name.Replace(" ", "_");

                if (UserResourceExtensions.PersistResourceValue(userCustomAction.TitleResource, $"CustomAction_{resourceKey}_Title", template, creationInfo))
                {
                    var customActionTitle = $"{{res:CustomAction_{resourceKey}_Title}}";
                    customAction.Title = customActionTitle;

                }
                if (UserResourceExtensions.PersistResourceValue(userCustomAction.DescriptionResource, $"CustomAction_{resourceKey}_Description", template, creationInfo))
                {
                    var customActionDescription = $"{{res:CustomAction_{resourceKey}_Description}}";
                    customAction.Description = customActionDescription;
                }
            }
#endif
            return customAction;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.CustomActions != null && (template.CustomActions.SiteCustomActions.Any() || template.CustomActions.WebCustomActions.Any());
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                var context = (ClientContext)web.Context;
                var webCustomActions = web.GetCustomActions();
                var siteCustomActions = context.Site.GetCustomActions();

                _willExtract = webCustomActions.Any() || siteCustomActions.Any();
            }
            return _willExtract.Value;
        }
    }
}
