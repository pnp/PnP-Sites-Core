using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeDevPnP.Core.Entities;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// JavaScript related methods
    /// </summary>
    public static partial class JavaScriptExtensions
    {
        /// <summary>
        /// Default Script Location value
        /// </summary>
        public const string SCRIPT_LOCATION = "ScriptLink";

        /// <summary>
        /// Injects links to javascript files via a adding a custom action to the site
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <param name="scriptLinks">semi colon delimited list of links to javascript files</param>
        /// <param name="sequence">Specifies the ordering priority for actions. A value is 0 indicates that the button will appear at the first position on the ribbon.</param>
        /// <returns>True if action was ok</returns>
        public static bool AddJsLink(this Web web, string key, string scriptLinks, int sequence = 0)
        {
            return AddJsLinkImplementation(web, key, new List<string>(scriptLinks.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)), sequence);
        }

        /// <summary>
        /// Injects links to javascript files via a adding a custom action to the site
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <param name="scriptLinks">semi colon delimited list of links to javascript files</param>
        /// <param name="sequence">Specifies the ordering priority for actions. A value is 0 indicates that the button will appear at the first position on the ribbon.</param>
        /// <returns>True if action was ok</returns>
        public static bool AddJsLink(this Site site, string key, string scriptLinks, int sequence = 0)
        {
            return AddJsLinkImplementation(site, key, new List<string>(scriptLinks.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)), sequence);
        }


        /// <summary>
        /// Injects links to javascript files via a adding a custom action to the site
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <param name="scriptLinks">IEnumerable list of links to javascript files</param>
        /// <param name="sequence">Specifies the ordering priority for actions. A value is 0 indicates that the button will appear at the first position on the ribbon.</param>
        /// <returns>True if action was ok</returns>
        public static bool AddJsLink(this Web web, string key, IEnumerable<string> scriptLinks, int sequence = 0)
        {
            return AddJsLinkImplementation(web, key, scriptLinks, sequence);
        }

        /// <summary>
        /// Injects links to javascript files via a adding a custom action to the site
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <param name="scriptLinks">IEnumerable list of links to javascript files</param>
        /// <param name="sequence">Specifies the ordering priority for actions. A value is 0 indicates that the button will appear at the first position on the ribbon.</param>
        /// <returns>True if action was ok</returns>
        public static bool AddJsLink(this Site site, string key, IEnumerable<string> scriptLinks, int sequence = 0)
        {
            return AddJsLinkImplementation(site, key, scriptLinks, sequence);
        }

        private static bool AddJsLinkImplementation(ClientObject clientObject, string key, IEnumerable<string> scriptLinks, int sequence)
        {
            bool ret;
            if (clientObject is Web || clientObject is Site)
            {
                var scriptLinksEnumerable = scriptLinks as string[] ?? scriptLinks.ToArray();
                if (!scriptLinksEnumerable.Any())
                {
                    throw new ArgumentException(nameof(scriptLinks));
                }

#if !ONPREMISES
                if (scriptLinksEnumerable.Length == 1)
                {
                    var scriptSrc = scriptLinksEnumerable[0];
                    if (!scriptSrc.StartsWith("http", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var serverUri = new Uri(clientObject.Context.Url);
                        if (scriptSrc.StartsWith("/"))
                        {
                            scriptSrc = $"{serverUri.Scheme}://{serverUri.Authority}{scriptSrc}";
                        }
                        else
                        {
                            var serverRelativeUrl = string.Empty;
                            if (clientObject is Web)
                            {
                                serverRelativeUrl = ((Web)clientObject).EnsureProperty(w => w.ServerRelativeUrl);
                            }
                            else
                            {
                                serverRelativeUrl = ((Site)clientObject).RootWeb.EnsureProperty(w => w.ServerRelativeUrl);
                            }
                            scriptSrc = $"{serverUri.Scheme}://{serverUri.Authority}{serverRelativeUrl}/{scriptSrc}";
                        }
                    }

                    var customAction = new CustomActionEntity
                    {
                        Name = key,
                        ScriptSrc = scriptSrc,
                        Location = SCRIPT_LOCATION,
                        Sequence = sequence
                    };
                    if (clientObject is Web)
                    {
                        ret = ((Web)clientObject).AddCustomAction(customAction);
                    }
                    else
                    {
                        ret = ((Site)clientObject).AddCustomAction(customAction);
                    }
                }
                else
                {
                    var scripts = new StringBuilder(@" var headID = document.getElementsByTagName('head')[0]; 
var scripts = document.getElementsByTagName('script');
var scriptsSrc = [];
for(var i = 0; i < scripts.length; i++) {
    if(scripts[i].type === 'text/javascript'){
        scriptsSrc.push(scripts[i].src);
    }
}
");
                    foreach (var link in scriptLinksEnumerable)
                    {
                        if (!string.IsNullOrEmpty(link))
                        {
                            scripts.Append(@"
if (scriptsSrc.indexOf('{1}') === -1)  {  
    var newScript = document.createElement('script');
    newScript.id = '{0}';
    newScript.type = 'text/javascript';
    newScript.src = '{1}';
    headID.appendChild(newScript);
    scriptsSrc.push('{1}');
}".Replace("{0}", key).Replace("{1}", link));
                        }

                    }

                    ret = AddJsBlockImplementation(clientObject, key, scripts.ToString(), sequence);
                }
#else
                var scripts = new StringBuilder(@" var headID = document.getElementsByTagName('head')[0]; 
var scripts = document.getElementsByTagName('script');
var scriptsSrc = [];
for(var i = 0; i < scripts.length; i++) {
    if(scripts[i].type === 'text/javascript'){
        scriptsSrc.push(scripts[i].src);
    }
}
");
                foreach (var link in scriptLinksEnumerable)
                {
                    if (!string.IsNullOrEmpty(link))
                    {
                        scripts.Append(@"
if (scriptsSrc.indexOf('{1}') === -1)  {  
    var newScript = document.createElement('script');
    newScript.id = '{0}';
    newScript.type = 'text/javascript';
    newScript.src = '{1}';
    headID.appendChild(newScript);
    scriptsSrc.push('{1}');
}".Replace("{0}", key).Replace("{1}", link));
                    }

                }

                ret = AddJsBlockImplementation(clientObject, key, scripts.ToString(), sequence);
#endif

            }
            else
            {
                throw new ArgumentException("Only Site or Web supported as clientObject");

            }
            return ret;
        }

        /// <summary>
        /// Removes the custom action that triggers the execution of a javascript link
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="key">Identifier (key) for the custom action that will be deleted</param>
        /// <returns>True if action was ok</returns>
        public static bool DeleteJsLink(this Web web, string key)
        {
            return DeleteJsLinkImplementation(web, key);
        }

        /// <summary>
        /// Removes the custom action that triggers the execution of a javascript link
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="key">Identifier (key) for the custom action that will be deleted</param>
        /// <returns>True if action was ok</returns>
        public static bool DeleteJsLink(this Site site, string key)
        {
            return DeleteJsLinkImplementation(site, key);
        }

        private static bool DeleteJsLinkImplementation(ClientObject clientObject, string key)
        {
            bool ret;
            if (clientObject is Web || clientObject is Site)
            {
                var jsAction = new CustomActionEntity()
                {
                    Name = key,
                    Location = SCRIPT_LOCATION,
                    Remove = true,
                };
                if (clientObject is Web)
                {
                    ret = ((Web)clientObject).AddCustomAction(jsAction);
                }
                else
                {
                    ret = ((Site)clientObject).AddCustomAction(jsAction);
                }

            }
            else
            {
                throw new ArgumentException("Only Site or Web supported as clientObject");
            }
            return ret;
        }

        /// <summary>
        /// Injects javascript via a adding a custom action to the site
        /// </summary>
        /// <param name="web">Site to be processed - can be root web or sub site</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <param name="scriptBlock">Javascript to be injected</param>
        /// <param name="sequence">Specifies the ordering priority for actions. A value is 0 indicates that the button will appear at the first position on the ribbon.</param>
        /// <returns>True if action was ok</returns>
        public static bool AddJsBlock(this Web web, string key, string scriptBlock, int sequence = 0)
        {
            return AddJsBlockImplementation(web, key, scriptBlock, sequence);

        }

        /// <summary>
        /// Injects javascript via a adding a custom action to the site
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <param name="scriptBlock">Javascript to be injected</param>
        /// <param name="sequence">Specifies the ordering priority for actions. A value is 0 indicates that the button will appear at the first position on the ribbon.</param>
        /// <returns>True if action was ok</returns>
        public static bool AddJsBlock(this Site site, string key, string scriptBlock, int sequence = 0)
        {
            return AddJsBlockImplementation(site, key, scriptBlock, sequence);
        }

        private static bool AddJsBlockImplementation(ClientObject clientObject, string key, string scriptBlock, int sequence)
        {
            bool ret;
            if (clientObject is Web || clientObject is Site)
            {
                var jsAction = new CustomActionEntity()
                {
                    Name = key,
                    Location = SCRIPT_LOCATION,
                    ScriptBlock = scriptBlock,
                    Sequence = sequence
                };
                if (clientObject is Web)
                {
                    ret = ((Web)clientObject).AddCustomAction(jsAction);
                }
                else
                {
                    ret = ((Site)clientObject).AddCustomAction(jsAction);
                }
            }
            else
            {
                throw new ArgumentException("Only Site or Web supported as clientObject");
            }
            return ret;
        }

        /// <summary>
        /// Checks if the target web already has a custom JsLink with a specified key
        /// </summary>
        /// <param name="web">Web to be processed</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <returns></returns>
        public static Boolean ExistsJsLink(this Web web, String key)
        {
            return (ExistsJsLinkImplementation(web, key));
        }

        /// <summary>
        /// Checks if the target site already has a custom JsLink with a specified key
        /// </summary>
        /// <param name="site">Site to be processed</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <returns>True if custom JsLink exists, false otherwise</returns>
        public static Boolean ExistsJsLink(this Site site, String key)
        {
            return (ExistsJsLinkImplementation(site, key));
        }

        /// <summary>
        /// Checks if the given clientObject already has a custom JsLink with a specified key
        /// </summary>
        /// <param name="clientObject">clientObject to operate on</param>
        /// <param name="key">Identifier (key) for the custom action that will be created</param>
        /// <returns>True if custom JsLink exists, false otherwise</returns>
        public static Boolean ExistsJsLinkImplementation(ClientObject clientObject, String key)
        {
            UserCustomActionCollection existingActions;
            if (clientObject is Web)
            {
                existingActions = ((Web)clientObject).UserCustomActions;
            }
            else
            {
                existingActions = ((Site)clientObject).UserCustomActions;
            }

            clientObject.Context.Load(existingActions);
            clientObject.Context.ExecuteQueryRetry();

            var actions = existingActions.ToArray();
            foreach (var action in actions)
            {
                if (action.Name == key &&
                    action.Location == "ScriptLink")
                {
                    return (true);
                }
            }

            return (false);
        }
    }
}

