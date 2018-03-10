#if !ONPREMISES
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Sites
{
    /// <summary>
    /// Class for communication site creation information
    /// </summary>
    public class CommunicationSiteCollectionCreationInformation
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
        /// The Guid of the site design to be used. If specified will override the SiteDesign property
        /// </summary>
        public Guid SiteDesignId { get; set; }

        /// <summary>
        /// The built-in site design to used. If both SiteDesignId and SiteDesign have been specified, the GUID specified as SiteDesignId will be used.
        /// </summary>
        public CommunicationSiteDesign SiteDesign { get; set; }

        /// <summary>
        /// If set to true, file sharing for guest users will be allowed.
        /// </summary>
        public bool AllowFileSharingForGuestUsers { get; set; }

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
        /// Default constructor
        /// </summary>
        public CommunicationSiteCollectionCreationInformation()
        {

        }

        /// <summary>
        /// CommunicationSiteCollectionCreationInformation constructor
        /// </summary>
        /// <param name="fullUrl">Url for the new communication site</param>
        /// <param name="title">Title of the site</param>
        /// <param name="description">Description of the site</param>
        public CommunicationSiteCollectionCreationInformation(string fullUrl, string title, string description = null)
        {
            this.Url = fullUrl;
            this.Title = title;
            this.Description = description;
        }
    }

    /// <summary>
    /// Class for site groupify information
    /// </summary>
    public class TeamSiteCollectionGroupifyInformation : SiteCreationInformation
    {

        /// <summary>
        /// If the site already has a modern home page, do we want to keep it?
        /// </summary>
        public bool KeepOldHomePage { get; set; }

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


    /// <summary>
    /// Class for site creation information
    /// </summary>
    public class TeamSiteCollectionCreationInformation : SiteCreationInformation
    {
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
    public abstract class SiteCreationInformation
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
        /// The Guid of the site design to be used. If specified will override the SiteDesign property
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
        /// </summary>
        public string Classification { get; set; }

        public SiteCreationInformation()
        {

        }

        public SiteCreationInformation(string alias, string displayName, string description = null)
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