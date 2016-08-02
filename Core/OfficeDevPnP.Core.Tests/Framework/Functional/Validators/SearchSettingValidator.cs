using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace OfficeDevPnP.Core.Tests.Framework.Functional.Validators
{

    public class SerializedSearchSetting
    {
        public string SchemaXml { get; set; }
    }

    public class SearchSettingValidator : ValidatorBase
    {
        #region construction        
        public SearchSettingValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }
        #endregion

        #region Validation logic
        public bool Validate(String sourceSearchSetting, string targetSearchSetting)
        {
            if (!String.IsNullOrEmpty(sourceSearchSetting))
            {
                if (String.IsNullOrEmpty(targetSearchSetting))
                {
                    return false;
                }
            }

            return true;
        }
        #endregion
    }
}
