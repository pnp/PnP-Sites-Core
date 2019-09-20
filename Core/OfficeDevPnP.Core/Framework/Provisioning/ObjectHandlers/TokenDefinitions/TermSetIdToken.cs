using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Attributes;
using System;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{termsetid:[groupname]:[termsetname]}",
      Description = "Returns the id of a term set given its name and its parent group",
      Example = "{termsetid:MyGroup:MyTermset}",
      Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
    internal class TermSetIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public TermSetIdToken(Web web, string groupName, string termsetName, Guid id)
            : base(web, $"{{termsetid:{Regex.Escape(groupName)}:{Regex.Escape(termsetName)}}}")
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