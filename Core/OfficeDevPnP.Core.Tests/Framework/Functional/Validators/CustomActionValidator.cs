using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System;
using System.Collections;
using System.Linq;
using System.Xml.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    public class CustomActionValidator: ValidatorBase
    {

        public static bool Validate(CustomActions sourceCustomActions, CustomActions targetCustomActions, TokenParser tokenParser, Web web)
        {

            if (web.IsNoScriptSite())
            {
                Console.WriteLine("Skipping validation of custom actions due to noscript site.");
                return true;
            }

            Console.WriteLine("Custom Action validation started...");

            bool isSiteCustomActionsMatch = false;
            bool isWebCustomActionsMatch = false;
            if (sourceCustomActions.SiteCustomActions.Count > 0)
            {
                isSiteCustomActionsMatch = ValidateCustomActions(sourceCustomActions.SiteCustomActions, targetCustomActions.SiteCustomActions, tokenParser, web);
                Console.WriteLine("Site Custom Actions validation " + isSiteCustomActionsMatch);
            }

            if (sourceCustomActions.WebCustomActions.Count > 0)
            {
                isWebCustomActionsMatch = ValidateCustomActions(sourceCustomActions.WebCustomActions, targetCustomActions.WebCustomActions, tokenParser, web);
                Console.WriteLine("Web Custom  Actions validation " + isWebCustomActionsMatch);
            }

            if (!isSiteCustomActionsMatch || !isWebCustomActionsMatch)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public static bool ValidateCustomActions(CustomActionCollection source, CustomActionCollection target, TokenParser tokenParser, Web web = null)
        {
            int sCount = 0;
            int tCount = 0;

            if (web != null && web.IsNoScriptSite())
            {
                Console.WriteLine("Skipping validation of custom actions due to noscript site.");
                return true;
            }

            foreach (CustomAction srcSCA in source)
            {
                //Only count the enabled ones
                if (srcSCA.Enabled && !srcSCA.Remove)
                {
                    // ensure token in source are parsed before comparing with target
                    srcSCA.Title = tokenParser.ParseString(srcSCA.Title);
                    srcSCA.ImageUrl = tokenParser.ParseString(srcSCA.ImageUrl);
                    srcSCA.ScriptBlock = tokenParser.ParseString(srcSCA.ScriptBlock);
                    srcSCA.ScriptSrc = tokenParser.ParseString(srcSCA.ScriptSrc, "~site", "~sitecollection");
                    srcSCA.Title = tokenParser.ParseString(srcSCA.Title);
                    srcSCA.Url = tokenParser.ParseString(srcSCA.Url);
                    if (srcSCA.CommandUIExtension != null)
                    {
                        srcSCA.CommandUIExtension = XElement.Parse(tokenParser.ParseString(srcSCA.CommandUIExtension.ToString()));
                    }

                    sCount++;
                    foreach (CustomAction tgtSCA in target)
                    {
                        if (tgtSCA.CommandUIExtension != null)
                        {
                            // Drop the namespace attribute before comparing (xmlns="http://schemas.microsoft.com/sharepoint"). 
                            // SharePoint injects this namespace when we extract a custom action that has a commandUIExtension
                            tgtSCA.CommandUIExtension = RemoveAllNamespaces(tgtSCA.CommandUIExtension);
                        }

                        // Use our custom action "Equals" implementation
                        if (srcSCA.Equals(tgtSCA))
                        {
                            tCount++;
                            break;
                        }
                        else
                        {
                            Console.WriteLine("{0} is not matching", tgtSCA.Name);
                        }
                    }
                }
            }

            if (sCount != tCount)
            {
                return false;
            }

            // cross check that enabled false custom actions do not exist anymore
            foreach (CustomAction srcSCA in source)
            {
                if (!srcSCA.Enabled || srcSCA.Remove)
                {
                    var ca = target.Where(w => w.Name == srcSCA.Name).FirstOrDefault();
                    if (ca != null)
                    {
                        return false;
                    }
                }
            }

            return true;
        }


        private static XElement RemoveAllNamespaces(XElement e)
        {
            return new XElement(e.Name.LocalName,
              (from n in e.Nodes()
               select ((n is XElement) ? RemoveAllNamespaces(n as XElement) : n)),
                  (e.HasAttributes) ?
                    (from a in e.Attributes()
                     where (!a.IsNamespaceDeclaration)
                     select new XAttribute(a.Name.LocalName, a.Value)) : null);
        }

    }

}
