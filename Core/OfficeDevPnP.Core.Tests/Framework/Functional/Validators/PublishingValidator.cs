using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Tests.Framework.Functional.Validators;
using System.Collections.Generic;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using Microsoft.SharePoint.Client;
using System.Linq;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    class PublishingValidator : ValidatorBase
    {
        public bool Validate(Publishing source, Publishing target, ClientContext clientContext)
        {

            if (clientContext.Web.IsNoScriptSite())
            {
                return true;
            }

            bool isAvailableWebTemplatesMatch = ValidateObjects(source.AvailableWebTemplates, target.AvailableWebTemplates, new List<string> { "LanguageCode", "TemplateName" });
            if (!isAvailableWebTemplatesMatch) { return false; }

            bool isPageLayoutsMatch = ValidateObjects(source.PageLayouts, target.PageLayouts, new List<string> { "IsDefault", "Path" });
            if (!isPageLayoutsMatch) { return false; }

            bool isValidDesignPackage = ValidateDesginPackage(source.DesignPackage, clientContext);
            if (!isValidDesignPackage) return false;

            return true;
        }

        private bool ValidateDesginPackage(DesignPackage SourceDesignPackage, ClientContext clientContext)
        {
            bool isValidDesignPackage = false;

            // Check if the solution file is uploaded
            var solutionGallery = clientContext.Site.RootWeb.GetCatalog((int)ListTemplateType.SolutionCatalog);
            var camlQuery = new CamlQuery();
            camlQuery.ViewXml = string.Format(@"<View>  
                                                    <Query> 
                                                        <Where><Eq><FieldRef Name='SolutionId' /><Value Type='Guid'>{0}</Value></Eq></Where> 
                                                    </Query> 
                                                        <ViewFields><FieldRef Name='ID' /></ViewFields> 
                                                </View>", SourceDesignPackage.PackageGuid);

            var solutions = solutionGallery.GetItems(camlQuery);
            clientContext.Load(solutions);
            clientContext.ExecuteQueryRetry();

            if (solutions.Count > 0)
            {
                string sourceesignPackageName = string.Format("{0}-v{1}.{2}.wsp", SourceDesignPackage.DesignPackagePath
                                                                                , SourceDesignPackage.MajorVersion
                                                                                , SourceDesignPackage.MinorVersion);

                foreach (ListItem packageItem in solutions)
                {
                    string targetdesignPackageName = Convert.ToString(packageItem["FileLeafRef"]);
                    
                    if (targetdesignPackageName.ToLower().Equals(sourceesignPackageName.ToLower()))
                    {
                        isValidDesignPackage = true;
                        break;
                    }
                }
            }

            return isValidDesignPackage;
        }
    }
}
