using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Handlers to be processed on the template
    /// </summary>
    [Flags]
    public enum Handlers : int
    {
        /// <summary>
        /// Value 1, represents AuditSettings
        /// </summary>
        AuditSettings = 1,
        /// <summary>
        /// Value 2, represents ComposedLook
        /// </summary>
        ComposedLook = 2,
        /// <summary>
        /// Value 4, represents CustomActions
        /// </summary>
        CustomActions = 4,
        /// <summary>
        /// Value 8, represents ExtensibilityProviders
        /// </summary>
        ExtensibilityProviders = 8,
        /// <summary>
        /// Value 16, represents Features
        /// </summary>
        Features = 16,
        /// <summary>
        /// Value 32, represents Fields
        /// </summary>
        Fields = 32,
        /// <summary>
        /// Value 64, represents Files
        /// </summary>
        Files = 64,
        /// <summary>
        /// Value 128, represents Lists
        /// </summary>
        Lists = 128,
        /// <summary>
        /// Value 256, represents Pages
        /// </summary>
        Pages = 256,
        /// <summary>
        /// Value 512, represents Publishing
        /// </summary>
        Publishing = 512,
        /// <summary>
        /// Value 1024, represents RegionalSettings
        /// </summary>
        RegionalSettings = 1024,
        /// <summary>
        /// Value 2048, represents SearchSettings
        /// </summary>
        SearchSettings = 2048,
        /// <summary>
        /// Value 4096, represents SitePolicy
        /// </summary>
        SitePolicy = 4096,
        /// <summary>
        /// Value 8192, represents SupportedUILanguages
        /// </summary>
        SupportedUILanguages = 8192,
        /// <summary>
        /// Value 16384, represents TermGroups
        /// </summary>
        TermGroups = 16384,
        /// <summary>
        /// Value 32768, represents Workflows
        /// </summary>
        Workflows = 32768,
        /// <summary>
        /// Value 65536, represents SiteSecurity
        /// </summary>
        SiteSecurity = 65536,
        /// <summary>
        /// Value 131072, represents ContentTypes
        /// </summary>
        ContentTypes = 131072,
        /// <summary>
        /// Value 262144, represents PropertyBagEntries
        /// </summary>
        PropertyBagEntries = 262144,
        /// <summary>
        /// Value 524288, represents PageContents
        /// </summary>
        PageContents = 524288,
        /// <summary>
        /// Value 1048576, represents WebSettings
        /// </summary>
        WebSettings = 1048576,
        /// <summary>
        /// Value 2097152, represents Navigation
        /// </summary>
        Navigation = 2097152,
        /// <summary>
        /// Value 4194304, represents Image Renditions
        /// </summary>
        ImageRenditions = 4194304,
        /// <summary>
        /// Value 8388608, represents Application Lifecycle Management
        /// </summary>
        ApplicationLifecycleManagement = 8388608,
        /// <summary>
        /// Value 16777216, represents Tenant
        /// </summary>
        Tenant = 16777216,
        /// <summary>
        /// Value 33554432, represents Web API Permissions
        /// </summary>
        WebApiPermissions = 33554432,
        /// <summary>
        /// Takes all handlers
        /// </summary>
        All = AuditSettings | ComposedLook | CustomActions | ExtensibilityProviders | Features | Fields | Files | Lists | Pages | Publishing | RegionalSettings | SearchSettings | SitePolicy | SupportedUILanguages | TermGroups | Workflows | SiteSecurity | ContentTypes | PropertyBagEntries | PageContents | WebSettings | Navigation | ImageRenditions | ApplicationLifecycleManagement | Tenant | WebApiPermissions
    }
}
