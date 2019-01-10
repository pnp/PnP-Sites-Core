using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{listid:[name]}",
     Description = "Returns a id of the list given its name",
     Example = "{listid:My List}",
     Returns = "f2cd6d5b-1391-480e-a3dc-7f7f96137382")]
    internal class ListIdToken : TokenDefinition
    {
        private string _listId = null;
        private string _name = null;
        public ListIdToken(Web web, string name, Guid listid)
            : base(web, $"{{listid:{Regex.Escape(name)}}}")
        {
            if (listid == Guid.Empty)
            {
                // on demand loading
                _name = name;
            }
            else
            {
                _listId = listid.ToString();
            }
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                if (_listId != null)
                {
                    CacheValue = _listId;
                }
                else
                {
                    List list = null;
#if !SP2013
                    try
                    {
#endif
                        list = TokenContext.Web.Lists.GetByTitle(_name);
                        TokenContext.Load(list, l => l.Id);
                        TokenContext.ExecuteQueryRetry();
#if !SP2013
                    }
                    catch (ServerException)
                    {
                        var mainLanguageName = GetListTitleForMainLanguage(TokenContext, _name);
                        list = TokenContext.Web.Lists.GetByTitle(mainLanguageName);
                        TokenContext.Load(list, l => l.Id);
                        TokenContext.ExecuteQueryRetry();
                    }
#endif
                    _listId = list.Id.ToString();
                    CacheValue = list.Id.ToString();
                }
            }
            return CacheValue;
        }

#if !SP2013
        private static Dictionary<String, String> listsTitles = new Dictionary<string, string>();

        /// <summary>
        /// This method retrieves the title of a list in the main language of the site
        /// </summary>
        /// <param name="context">The current ClientContext</param>
        /// <param name="name">The title of the list in the current user's language</param>
        /// <returns>The title of the list in the main language of the site</returns>
        private String GetListTitleForMainLanguage(ClientContext context, String name)
        {
            if (listsTitles.ContainsKey(name))
            {
                // Return the title that we already have
                return (listsTitles[name]);
            }
            else
            {
                // Refresh the list of titles and get the main language
                context.Load(context.Web.Lists, ls => ls.Include(l => l.Id, l => l.Title, l => l.TitleResource));
                context.Load(context.Web, w => w.Language);
                context.ExecuteQueryRetry();

                // Get the default culture for the current web
                var ci = new System.Globalization.CultureInfo((int)context.Web.Language);

                // Refresh the list of lists with a lock
                lock (typeof(ListIdToken))
                {
                    // Reset the cache of lists titles
                    ListIdToken.listsTitles.Clear();

                    // Add the new lists title using the main language of the site
                    foreach (var list in context.Web.Lists)
                    {
                        var titleResource = list.TitleResource.GetValueForUICulture(ci.Name);
                        context.ExecuteQueryRetry();
                        ListIdToken.listsTitles.Add(list.Title, titleResource.Value);
                    }
                }

                // If now we have the list title ...
                if (listsTitles.ContainsKey(name))
                {
                    // Return the title, if any
                    return (listsTitles[name]);
                }
                else
                {
                    return (null);
                }
            }
        }
    }
#endif
}