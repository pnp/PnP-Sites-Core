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
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Utilities;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class FilesValidator : ValidatorBase
    {
        public FilesValidator() : base()
        {
            // optionally override schema version
            //SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(Core.Framework.Provisioning.Model.FileCollection sourceFiles, Microsoft.SharePoint.Client.ClientContext ctx)
        {
            int scount = 0;
            int tcount = 0;

            try
            {
                // Check if this is not a noscript site as we're not allowed to write to the web property bag is that one
                bool isNoScriptSite = ctx.Web.IsNoScriptSite();

                foreach (var sf in sourceFiles)
                {
                    scount++;
                    string fileName = sf.Src;
                    string folderName = sf.Folder;
                    string fileUrl = UrlUtility.Combine(ctx.Web.ServerRelativeUrl, folderName + "/" + fileName);

                    // Skip the files we skipped to provision (if any)
                    if (ObjectFiles.SkipFile(isNoScriptSite, fileName, folderName))
                    {
                        continue;
                    }

                    var file = ctx.Web.GetFileByServerRelativeUrl(UrlUtility.Combine(ctx.Web.ServerRelativeUrl, folderName + "/" + fileName));
                    ctx.Load(file, f => f.Exists, f => f.Length);
                    ctx.ExecuteQueryRetry();

                    if (file.Exists)
                    {
                        tcount++;

                        #region File - Security
                        if (sf.Security != null)
                        {
                            ctx.Load(file, f => f.ListItemAllFields);
                            ctx.ExecuteQueryRetry();
                            bool isSecurityMatch = ValidateSecurityCSOM(ctx, sf.Security, file.ListItemAllFields);
                            if (!isSecurityMatch)
                            {
                                return false;
                            }

                        }
                        #endregion

                        #region Overwrite validation
                        if (sf.Overwrite == false)
                        {
                            // lookup the original added file size...should be different from the one we retrieved from SharePoint since we opted to NOT overwrite
                            var files = System.IO.Directory.GetFiles(@".\framework\functional\templates");
                            foreach (var f in files)
                            {
                                if (f.Contains(sf.Src))
                                {
                                    if (new System.IO.FileInfo(f).Length == file.Length)
                                    {
                                        return false;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch(Exception ex)
            {
                // Return false if we get an exception
                Console.WriteLine(ex.ToDetailedString(ctx));
                return false;
            }

            return true;
        }

        public bool Validate1605(ProvisioningTemplate template, Microsoft.SharePoint.Client.ClientContext ctx)
        {
            var directoryFiles = new List<Core.Framework.Provisioning.Model.File>();
            
            // Get all files from directories
            foreach (var directory in template.Directories)
            {
                var metadataProperties = directory.GetMetadataProperties();
                directoryFiles = directory.GetDirectoryFiles(metadataProperties);

                // Add directory files to template file collection
                foreach (var dFile in directoryFiles)
                {
                    var file = new Core.Framework.Provisioning.Model.File();
                    file.Src = dFile.Src.Replace(directory.Src + "\\", "");
                    file.Folder = directory.Folder;
                    template.Files.Add(file);
                }
            }
           
            // validate all files
            return Validate(template.Files, ctx);
        }
    }
}
