using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using System.Xml;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Linq;
using OfficeDevPnP.Core.Utilities;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class PagesValidator : ValidatorBase
    {
        private static string[] chromeType = new string[] { "NA", "Default", "None", "Title and Border", "Title Only", "Border Only" };
        private static string[] chromeState = new string[] { "Normal", "Minimized" };
        private static string[] direction = new string[] { "NotSet", "Left to Right", "Right to Left" };
        private static string[] exportMode = new string[] { "Non-sensitive data only", "Export all data" };
        private static string[] helpMode = new string[] { "Modal", "Modeless", "Navigate" };
        public PagesValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(PageCollection sourcePages, Microsoft.SharePoint.Client.ClientContext ctx)
        {
            int scount = 0;
            int tcount = 0;


            Web web = ctx.Web;
            ctx.Load(web, w => w.Url, w => w.ServerRelativeUrl);
            ctx.ExecuteQueryRetry();

            foreach (var sourcePage in sourcePages)
            {

                string pageUrl = sourcePage.Url.ToString();
                pageUrl = pageUrl.Replace("{site}", web.ServerRelativeUrl);

                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(pageUrl);
                ctx.Load(file, page => page.ListItemAllFields, page => page.ListItemAllFields.RoleAssignments.Include(roleAsg => roleAsg.Member,
                  roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));
                ctx.ExecuteQueryRetry();

                if (file != null)
                {
                    #region Page - Fields

                    if (sourcePage.Fields.Count > 0)
                    {
                        scount = 0;
                        tcount = 0;
                        string sourceWikifield = sourcePage.Fields["WikiField"].ToString();
                        string targetwikiField = (string)file.ListItemAllFields["WikiField"];

                        if (sourceWikifield.Trim() != RemoveHTMLTags(targetwikiField).Trim())
                        {
                            return false;
                        }
                    }

                    #endregion

                    #region  Page - Webparts

                    if (!ctx.Web.IsNoScriptSite() && sourcePage.WebParts.Count > 0)
                    {
                        LimitedWebPartManager wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
                        ctx.Load(wpm.WebParts, wps => wps.Include(wp => wp.WebPart.Title, wp => wp.WebPart.Properties));
                        ctx.ExecuteQueryRetry();

                        if (wpm.WebParts.Count > 0)
                        {
                            foreach (var spwp in sourcePage.WebParts)
                            {
                                scount++;
                                foreach (WebPartDefinition wpd in wpm.WebParts)
                                {
                                    if (spwp.Title == wpd.WebPart.Title)
                                    {
                                        tcount++;

                                        //Page - Webpart Properties 
                                        Microsoft.SharePoint.Client.WebParts.WebPart wp = wpd.WebPart;
                                        var isWebPropertiesMatch = CSOMWebPartPropertiesValidation(spwp.Contents, wp.Properties);
                                        if (!isWebPropertiesMatch)
                                        {
                                            return false;
                                        }
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                            if (scount != tcount)
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }

                    #endregion

                    #region Page - Security

                    scount = 0;
                    tcount = 0;

                    if (sourcePage.Security != null && file.ListItemAllFields.RoleAssignments.Count > 0)
                    {
                        bool securityResult = ValidateSecurityCSOM(ctx, sourcePage.Security, file.ListItemAllFields);

                        if (!securityResult)
                        {
                            return false;
                        }
                    }
                    #endregion
                }
                else
                {
                    return false;
                }
            }

            return true;
        }
        private static bool CSOMWebPartPropertiesValidation(string sourceContent, PropertyValues properties)
        {
            int scount = 0;
            int tcount = 0;
            XmlDocument sDoc = new XmlDocument();
            sDoc.LoadXml(sourceContent);
            XmlNodeList sProps = sDoc.GetElementsByTagName("property");
            string[] ignoreProperties = { "xmldefinition", "height", "width" };

            if (sProps.Count != 0)
            {
                foreach (XmlNode xnSource in sProps)
                {
                    string saName = xnSource.Attributes["name"].InnerText.ToString();
                    string saType = xnSource.Attributes["type"].InnerText.ToString();
                    string saText = xnSource.InnerText;

                    if (!ignoreProperties.Contains(saName.ToLower()))
                    {
                        scount++;

                        foreach (var tVariable in properties.FieldValues)
                        {
                            if (saName == tVariable.Key && saText == Convert.ToString(tVariable.Value))
                            {
                                tcount++;
                                break;
                            }
                            else if (saName == tVariable.Key)
                            {
                                int position = Convert.ToInt32(tVariable.Value);
                                string sourceText = saText.ToLower().Trim();

                                bool result = CheckProperty(tVariable.Key, sourceText, position);

                                if (result)
                                {
                                    tcount++;
                                }
                                else
                                {
                                    return false;
                                }

                            }
                        }
                    }
                }
            }
            else if (sDoc.LastChild.Name == "WebPart")
            {
                sProps = sDoc.LastChild.ChildNodes;

                foreach (XmlNode sProp in sProps)
                {
                    string speName = sProp.Name;
                    string speText = sProp.InnerText;
                    if (speName.ToLower() != "height" && speName.ToLower() != "width")
                    {
                        scount++;

                        foreach (var pfvalue in properties.FieldValues)
                        {
                            if (speName == pfvalue.Key && speText == Convert.ToString(pfvalue.Value))
                            {
                                tcount++;
                                break;
                            }
                            else if (speName == pfvalue.Key)
                            {
                                int position = Convert.ToInt32(pfvalue.Value);
                                string sourceText = speName.ToLower().Trim();

                                bool result = CheckProperty(pfvalue.Key, sourceText, position);

                                if (result)
                                {
                                    tcount++;
                                    break;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                        }
                    }
                }
            }
            if (scount != tcount)
            {
                return false;
            }

            return true;
        }
        private static bool CheckProperty(string key, string sourceText, int position)
        {
            bool isExists = false;
            switch (key)
            {
                case "ChromeState":
                    if (sourceText == chromeState[position].ToString().ToLower())
                    {
                        isExists = true;
                    }
                    else
                    {
                        //Console.WriteLine("---- CSOMWebPartPropertiesValidation - ChromeState webpart property of source and target are not matching");
                    }
                    break;
                case "ChromeType":

                    if (sourceText == chromeType[position].ToString().ToLower())
                    {
                        isExists = true;
                    }
                    else
                    {
                        //Console.WriteLine("---- CSOMWebPartPropertiesValidation - ChromeType webpart property of source and target are not matching");
                    }
                    break;
                case "Direction":

                    if (sourceText == direction[position].ToString().ToLower())
                    {
                        isExists = true;
                    }
                    else
                    {
                        //Console.WriteLine("---- CSOMWebPartPropertiesValidation - Direction webpart property of source and target are not matching");
                    }
                    break;
                case "ExportMode":

                    if (exportMode[position].ToString().ToLower().Contains(sourceText))
                    {
                        isExists = true;
                    }
                    else
                    {
                        //Console.WriteLine("---- CSOMWebPartPropertiesValidation - ChromeState ExportMode property of source and target are not matching");
                    }
                    break;

                case "HelpMode":

                    if (sourceText == helpMode[position].ToString().ToLower())
                    {
                        isExists = true;
                    }
                    else
                    {
                        //Console.WriteLine("---- CSOMWebPartPropertiesValidation - HelpMode webpart property of source and target are not matching");
                    }
                    break;
                default: break;
            }

            return isExists;
        }
        private static string RemoveHTMLTags(string source)
        {
            char[] array = new char[source.Length];
            int arrayIndex = 0;
            bool inside = false;

            for (int i = 0; i < source.Length; i++)
            {
                char let = source[i];
                if (let == '<')
                {
                    inside = true;
                    continue;
                }
                if (let == '>')
                {
                    inside = false;
                    continue;
                }
                if (!inside)
                {
                    array[arrayIndex] = let;
                    arrayIndex++;
                }
            }
            return new string(array, 0, arrayIndex);
        }
    }
}
