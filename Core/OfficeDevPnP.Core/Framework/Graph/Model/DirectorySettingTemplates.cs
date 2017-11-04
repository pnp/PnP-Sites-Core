using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Defines the container for a collection of DirectorySetting objects
    /// </summary>
    public class DirectorySettingTemplates
    {
        [JsonProperty(PropertyName = "value")]
        public List<DirectorySetting> Templates { get; set; }
    }
}
