using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Represents a DirectorySetting for Azure AD
    /// </summary>
    public class DirectorySetting
    {
        public Guid Id { get; set; }

        public DateTime? DeletedDateTime { get; set; }

        public String Description { get; set; }

        public String DisplayName { get; set; }

        [JsonProperty(PropertyName="values")]
        public List<DirectorySettingValue> SettingValues { get; set; }
    }
}
