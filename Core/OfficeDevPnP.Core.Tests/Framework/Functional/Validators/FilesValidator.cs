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

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class FilesValidator : ValidatorBase
    {
        public FilesValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(FileCollection sourceFiles, Microsoft.SharePoint.Client.ClientContext ctx)
        {
            int scount = 0;
            int tcount = 0;

            foreach (var sf in sourceFiles)
            {
                scount++;
                string fileName = sf.Src;
                string folderName = sf.Folder;
                string fileUrl = UrlUtility.Combine(ctx.Web.ServerRelativeUrl, folderName + "/" + fileName);
                var file = ctx.Web.GetFileByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, folderName + "/" + fileName));
                ctx.Load(file, f => f.Exists);
                ctx.ExecuteQuery();

                if (file.Exists)
                {
                    tcount++;

                    #region File - Security
                    if (sf.Security != null)
                    {
                        ctx.Load(file, f => f.ListItemAllFields);
                        ctx.ExecuteQuery();
                        bool isSecurityMatch = ValidateSecurityCSOM(ctx, sf.Security, file.ListItemAllFields);
                        if (!isSecurityMatch)
                        {
                            return false;
                        }

                    }
                    #endregion
                    #region Webparts validation
                    if (sf.WebParts.Count > 0)
                    {
                        WebpartValidator wpv = new WebpartValidator();
                        bool isWepartMatch = wpv.Validate(ctx, sf, file);
                        if (!isWepartMatch)
                        {
                            return false;
                        }
                    }
                    #endregion
                }
            }
            return true;
        }
    }
}
