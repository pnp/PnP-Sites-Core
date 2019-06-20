using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectiontermsetid:[termsetname]}",
        Description = "Returns the id of the given termset name located in the sitecollection termgroup",
        Example = "{sitecollectiontermsetid:MyTermset}",
        Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
    internal class SiteCollectionTermSetIdToken : TokenDefinition
    {
        private string _value;

        public SiteCollectionTermSetIdToken(Web web, string termsetName, Guid id)
            : base(web, $"{{sitecollectiontermsetid:{Regex.Escape(termsetName)}}}")
        {
            _value = id.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}