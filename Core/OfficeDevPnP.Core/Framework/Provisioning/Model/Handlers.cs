using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    [Flags]
    public enum Handlers : int
    {
        AuditSettings = 1,
        ComposedLook = 2,
        CustomActions = 4,
        ExtensibilityProviders = 8,
        Features = 16,
        Fields = 32,
        Files = 64,
        Lists = 128,
        Pages = 256,
        Publishing = 512,
        RegionalSettings = 1024,
        SearchSettings = 2048,
        SitePolicy = 4096,
        SupportedUILanguages = 8192,
        TermGroups = 16384,
        Workflows = 32768,
        SiteSecurity = 65536,
        ContentTypes = 131072,
        PropertyBagEntries = 262144,
        PageContents = 524288,
        WebSettings = 1048576,
        All = AuditSettings | ComposedLook | CustomActions | ExtensibilityProviders | Features | Fields | Files | Lists | Pages | Publishing | RegionalSettings | SearchSettings | SitePolicy | SupportedUILanguages | TermGroups | Workflows | SiteSecurity | ContentTypes | PropertyBagEntries | PageContents | WebSettings
    }
}
