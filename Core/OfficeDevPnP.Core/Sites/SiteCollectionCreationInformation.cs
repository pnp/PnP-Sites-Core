#if !SP2013 && !SP2016
using System;

namespace OfficeDevPnP.Core.Sites
{
    /// <summary>
    /// Class for communication site creation information
    /// </summary>
    public class CommunicationSiteCollectionCreationInformation : SiteCreationInformation
    {
        /// <summary>
        /// The Guid of the site design to be used. If specified will override the SiteDesign property
        /// </summary>
        public Guid SiteDesignId { get; set; }

        /// <summary>
        /// The built-in site design to used. If both SiteDesignId and SiteDesign have been specified, the GUID specified as SiteDesignId will be used.
        /// </summary>
        public CommunicationSiteDesign SiteDesign { get; set; }

        /// <summary>
        /// The Guid of the hub site to be used. If specified will associate the communication site to the hub site
        /// </summary>
        public Guid HubSiteId { get; set; }
        
        /// <summary>
        /// The Sensitivity label to use. For instance 'Top Secret'. See https://www.youtube.com/watch?v=NxvUXBiPFcw for more information.
        /// </summary>
        public string SensitivityLabel { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public CommunicationSiteCollectionCreationInformation() : this(string.Empty, string.Empty)
        {
        }

        /// <summary>
        /// CommunicationSiteCollectionCreationInformation constructor
        /// </summary>
        /// <param name="fullUrl">Url for the new communication site</param>
        /// <param name="title">Title of the site</param>
        /// <param name="description">Description of the site</param>
        public CommunicationSiteCollectionCreationInformation(string fullUrl, string title, string description = null) : base(fullUrl, title, description)
        {
            WebTemplate = "SITEPAGEPUBLISHING#0";
        }
    }

    /// <summary>
    /// Class for Team site with no group creation information
    /// </summary>
    public class TeamNoGroupSiteCollectionCreationInformation : SiteCreationInformation
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public TeamNoGroupSiteCollectionCreationInformation() : this(string.Empty, string.Empty)
        {
        }

        /// <summary>
        /// TeamNoGroupSiteCollectionCreationInformation constructor
        /// </summary>
        /// <param name="fullUrl">Url for the new team site</param>
        /// <param name="title">Title of the site</param>
        /// <param name="description">Description of the site</param>
        public TeamNoGroupSiteCollectionCreationInformation(string fullUrl, string title, string description = null) : base(fullUrl, title, description)
        {
            WebTemplate = "STS#3";
        }
    }

    /// <summary>
    /// Class for site creation information
    /// </summary>
    public abstract class SiteCreationInformation
    {
        /// <summary>
        /// The fully qualified URL (e.g. https://yourtenant.sharepoint.com/sites/mysitecollection) of the site.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// The title of the site to create
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The owner of the site. Reserved for future use.
        /// </summary>
        public string Owner { get; set; }

        /// <summary>
        /// If set to true, file sharing for guest users will be allowed.
        /// </summary>
        [Obsolete("This property is obsolete, use ShareByEmailEnabled instead")]
        public bool AllowFileSharingForGuestUsers
        {
            get
            {
                return ShareByEmailEnabled;
            }
            set
            {
                ShareByEmailEnabled = value;
            }
        }

        /// <summary>
        /// If set to true sharing files by email is enabled. Defaults to false.
        /// </summary>
        public bool ShareByEmailEnabled { get; set; }

        /// <summary>
        /// The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
        /// </summary>
        public string Classification { get; set; }

        /// <summary>
        /// The description to use for the site.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The language to use for the site. If not specified will default to the language setting of the clientcontext.
        /// </summary>
        public uint Lcid { get; set; }

        /// <summary>
        /// The Web template to use for the site.
        /// </summary>
        public string WebTemplate { get; protected set; }

        /// <summary>
        /// The geography in which to create the site collection. Only applicable to multi-geo enabled tenants.
        /// </summary>
        public Enums.Office365Geography? PreferredDataLocation { get; set; }

        public SiteCreationInformation()
        {
        }

        public SiteCreationInformation(string fullUrl, string title, string description = null)
        {
            this.Url = fullUrl;
            this.Title = title;
            this.Description = description;
        }
    }

#if !ONPREMISES
    /// <summary>
    /// Class for site groupify information
    /// </summary>
    public class TeamSiteCollectionGroupifyInformation : SiteCreationGroupInformation
    {

        /// <summary>
        /// If the site already has a modern home page, do we want to keep it?
        /// </summary>
        public bool KeepOldHomePage { get; set; }

        /// <summary>
        /// Set the owners of the modern team site. Specify the UPN values in a string array.
        /// </summary>
        public string[] Owners { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public TeamSiteCollectionGroupifyInformation() : base()
        {
        }

        /// <summary>
        /// TeamSiteCollectionTeamSiteCollectionGroupifyInformationCreationInformation constructor
        /// </summary>
        /// <param name="alias">Alias for the group which will be linked to this site</param>
        /// <param name="displayName">Name of the site</param>
        /// <param name="description">Title of the site</param>
        public TeamSiteCollectionGroupifyInformation(string alias, string displayName, string description = null) : base(alias, displayName, description)
        {
        }
    }
#endif
    /// <summary>
    /// Class for group site creation information
    /// </summary>
    public class TeamSiteCollectionCreationInformation : SiteCreationGroupInformation
    {
        /// <summary>
        /// Set the owners of the modern team site. Specify the UPN values in a string array.
        /// </summary>
        public string[] Owners { get; set; }

        /// <summary>
        /// The ID of the Site Design to apply, if any
        /// </summary>
        public Guid? SiteDesignId { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public TeamSiteCollectionCreationInformation() : base()
        {
        }

        /// <summary>
        /// TeamSiteCollectionCreationInformation constructor
        /// </summary>
        /// <param name="alias">Alias for the group linked to this site</param>
        /// <param name="displayName">Name of the site</param>
        /// <param name="description">Title of the site</param>
        public TeamSiteCollectionCreationInformation(string alias, string displayName, string description = null) : base(alias, displayName, description)
        {
        }
    }

    /// <summary>
    /// Base class for site creation/groupify information
    /// </summary>
    public abstract class SiteCreationGroupInformation
    {
        //{"displayName":"test modernteamsite","alias":"testmodernteamsite","isPublic":true,"optionalParams":{"Description":"","CreationOptions":{"results":[]},"Classification":""}}

        /// <summary>
        /// Alias of the underlying Office 365 Group
        /// </summary>
        public string Alias { get; set; }

        /// <summary>
        /// The title of the site to create
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Defines whether the Office 365 Group will be public (default), or private.
        /// </summary>
        public bool IsPublic { get; set; } = true;

        /// <summary>
        /// The description of the site to be created.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
        /// </summary>
        public string Classification { get; set; }

        public uint Lcid { get; set; }

        /// <summary>
        /// The Guid of the hub site to be used. If specified will associate the modern team site to the hub site.
        /// </summary>
        public Guid HubSiteId { get; set; }

        /// <summary>
        /// The Sensitivity label to use. For instance 'Top Secret'. See https://www.youtube.com/watch?v=NxvUXBiPFcw for more information.
        /// </summary>
        public string SensitivityLabel { get; set; }

        /// <summary>
        /// The geography in which to create the site collection. Only applicable to multi-geo enabled tenants.
        /// </summary>
        public Enums.Office365Geography? PreferredDataLocation { get; set; }

        public SiteCreationGroupInformation()
        {

        }

        public SiteCreationGroupInformation(string alias, string displayName, string description = null)
        {
            this.Alias = alias;
            this.DisplayName = displayName;
            this.Description = description;
        }
    }

    public enum CommunicationSiteDesign
    {
        Topic = 0,
        Showcase = 1,
        Blank = 2,
    }

}
#endif