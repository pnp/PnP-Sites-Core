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
        PublishingPages = 512,
        Publishing = 1024,
        RegionalSettings = 2048,
        SearchSettings = 4096,
        SitePolicy = 8192,
        SupportedUILanguages = 16384,
        TermGroups = 32768,
        Workflows = 65536,
        SiteSecurity = 131072,
        ContentTypes = 262144,
        PropertyBagEntries = 524288,
        PageContents = 1048576,
        WebSettings = 2097152,
        All = AuditSettings | ComposedLook | CustomActions | ExtensibilityProviders | Features | Fields | Files | Lists | Pages | PublishingPages | Publishing | RegionalSettings | SearchSettings | SitePolicy | SupportedUILanguages | TermGroups | Workflows | SiteSecurity | ContentTypes | PropertyBagEntries | PageContents | WebSettings
    }
}
